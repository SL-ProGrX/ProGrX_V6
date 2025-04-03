VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_RepMovTipoDocumento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Documentos"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9165
   HelpContextID   =   22
   Icon            =   "frmCntX_RepMovTipoDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   8772
      _Version        =   1310723
      _ExtentX        =   15473
      _ExtentY        =   2138
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_RepMovTipoDocumento.frx":000C
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4812
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1920
      TabIndex        =   7
      Top             =   3240
      Width           =   1332
      _Version        =   1310723
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1920
      TabIndex        =   8
      Top             =   3600
      Width           =   1332
      _Version        =   1310723
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
   Begin XtremeSuiteControls.GroupBox fraRangos 
      Height          =   1572
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   8532
      _Version        =   1310723
      _ExtentX        =   15049
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Tipos de Documentos y Cuentas:"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtDesde 
         Height          =   312
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtHasta 
         Height          =   312
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDesdeDesc 
         Height          =   312
         Left            =   3000
         TabIndex        =   12
         Top             =   720
         Width           =   4812
         _Version        =   1310723
         _ExtentX        =   8488
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtHastaDesc 
         Height          =   312
         Left            =   3000
         TabIndex        =   13
         Top             =   1080
         Width           =   4812
         _Version        =   1310723
         _ExtentX        =   8488
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTipo 
         Height          =   312
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoDesc 
         Height          =   312
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   4812
         _Version        =   1310723
         _ExtentX        =   8488
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   720
         Width           =   612
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   612
      End
   End
   Begin XtremeSuiteControls.CheckBox chkRango 
      Height          =   252
      Left            =   5040
      TabIndex        =   20
      Top             =   1320
      Width           =   3612
      _Version        =   1310723
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Filtrar el rango de cuentas y documentos"
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
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos (Asientos Registrados)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   1875
      TabIndex        =   19
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   612
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Width           =   612
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Informe: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   4560
      TabIndex        =   3
      Top             =   3600
      Width           =   2412
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   9252
   End
End
Attribute VB_Name = "frmCntX_RepMovTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBusca As Integer

Private Function fxDescAsiento(vTipo As String, vNumAsiento As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from Cntx_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & vTipo & "' and num_asiento = '" & vNumAsiento & "'"
rsX.CursorLocation = adUseServer
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxDescAsiento = ""
Else
 fxDescAsiento = rsX!Descripcion
End If
rsX.Close

End Function


Private Sub chkRango_Click()
If chkRango.Value = 0 Then
  fraRangos.Enabled = False
Else
  fraRangos.Enabled = True
  txtTipo = ""
  txtDesde = ""
  txtHasta = ""
  txtTipo.SetFocus
End If
End Sub


Private Sub sbDesbalance()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

lbl.Caption = "Rastreando Desbalances [Espere]"
lbl.Refresh

strSQL = "select A.cod_contabilidad,A.tipo_asiento,A.num_asiento" _
       & " from Cntx_Asientos A inner join Cntx_Asientos_detalle D on A.cod_contabilidad = D.cod_contabilidad" _
       & " and A.tipo_asiento = D.tipo_asiento and A.num_asiento = D.num_asiento" _
       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and A.fecha_asiento between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'" _
       & " group by A.cod_contabilidad,A.tipo_asiento,A.num_asiento" _
       & " Having (Sum(D.monto_debito) - Sum(D.monto_credito)) <> 0"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 strSQL = "update Cntx_Asientos set balanceado = 'N' where cod_contabilidad = " & rs!COD_CONTABILIDAD _
        & " and tipo_asiento = '" & rs!Tipo_Asiento & "' and num_asiento = '" _
        & rs!Num_Asiento & "'"
 Call ConectionExecute(strSQL, 0)
 rs.MoveNext
Loop
rs.Close

lbl.Caption = ""

Me.MousePointer = vbDefault

End Sub


Private Sub cmdReporte_Click()
Dim strSQL As String, strSubtitulo As String

'Verifica

If dtpInicio.Value > dtpCorte.Value Then
  MsgBox "Rango de fecha no es válido, Verifique...", vbCritical
  Exit Sub
End If


If chkRango.Value = 1 Then
    Select Case ""
      Case txtDesde, txtHasta, txtTipo
        MsgBox "Datos especificados en los rangos, no son válidos. Verifiquelos...", vbCritical
        Exit Sub
    End Select
End If

strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
       & " AND {Cntx_Asientos.FECHA_ASIENTO} in Date(" & Year(dtpInicio.Value) & "," _
       & Month(dtpInicio.Value) & "," & Day(dtpInicio.Value) _
       & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) & "," _
       & Day(dtpCorte.Value) & ")"

strSubtitulo = "Inicio: " & dtpInicio.Value & " Corte: " & dtpCorte.Value _
             & " Asientos: " & Mid(cbo.Text, 6, 30)

If chkRango.Value = 1 Then
   strSQL = strSQL & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtTipo & "' AND " _
          & "{Cntx_Asientos.NUM_ASIENTO} >= '" & txtDesde & "' AND {Cntx_Asientos.NUM_ASIENTO}" _
          & " <= '" & txtHasta & "'"
   strSubtitulo = strSubtitulo & " [TIPO: " & txtTipo & " Desde: " & txtDesde _
                & " Hasta: " & txtHasta & "] Asientos: " & Mid(cbo.Text, 6, 30)
End If

Select Case Mid(cbo, 1, 2)
    Case "02" 'Aplicados
       strSQL = strSQL & " AND ISNULL({Cntx_Asientos.FECHA_APLICADO}) = FALSE"
    Case "03" 'Sin Aplicar
       strSQL = strSQL & " AND ISNULL({Cntx_Asientos.FECHA_APLICADO}) = TRUE"
    Case "04" 'Desbalanceados
       Call sbDesbalance
       strSQL = strSQL & " AND {Cntx_Asientos.BALANCEADO} = 'N'"
    Case Else
      'Todos
End Select


Call sbCntX_Reportes("Asientos", strSQL, strSubtitulo)

End Sub

Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


dtpInicio.Value = Date
dtpCorte.Value = dtpInicio.Value

cbo.Clear
cbo.AddItem "01 - Todos"
cbo.AddItem "02 - Aplicados"
cbo.AddItem "03 - Sin Aplicar"
cbo.AddItem "04 - Desbalanceados"

cbo.Text = "01 - Todos"


vBusca = 1
End Sub

Private Sub sbConsultas()
Select Case vBusca
  Case 1 'Tipo de ASiento
     gBusquedas.Columna = "Tipo_Asiento"
     gBusquedas.Orden = "Tipo_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
     frmBusquedas.Show vbModal
     txtTipo = gBusquedas.Resultado
  Case 2 'Numero de Asiento Desde
     gBusquedas.Columna = "Num_Asiento"
     gBusquedas.Orden = "Num_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and tipo_asiento = '" & txtTipo & "' and " _
                       & " anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " & gCntX_Parametros.PeriodoMes
     gBusquedas.Consulta = "select Num_asiento,descripcion from Cntx_Asientos"
     frmBusquedas.Show vbModal
     txtDesde = gBusquedas.Resultado
     
  Case 3 'Numero de Asiento Hasta
     gBusquedas.Columna = "Num_Asiento"
     gBusquedas.Orden = "Num_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and tipo_asiento = '" & txtTipo & "' and " _
                       & " anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " & gCntX_Parametros.PeriodoMes
     gBusquedas.Consulta = "select Num_asiento,descripcion from Cntx_Asientos"
     frmBusquedas.Show vbModal
     txtHasta = gBusquedas.Resultado

End Select


End Sub


Private Sub txtDesde_GotFocus()
vBusca = 2
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtHasta.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub txtDesde_LostFocus()
If txtTipo <> "" And txtDesde <> "" Then
  txtDesdeDesc.Text = fxDescAsiento(txtTipo.Text, txtDesde.Text)
End If
End Sub

Private Sub txtHasta_GotFocus()
vBusca = 3
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub txtHasta_LostFocus()
If txtTipo <> "" And txtHasta <> "" Then
  txtHastaDesc.Text = fxDescAsiento(txtTipo.Text, txtHasta.Text)
End If
End Sub

Private Sub txtTipo_GotFocus()
vBusca = 1
End Sub

Private Sub txtTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesde.SetFocus
If KeyCode = vbKeyF4 Then sbConsultas
End Sub

Private Sub txtTipo_LostFocus()
txtTipoDesc.Text = fxCntX_TiposAsientos("D", txtTipo.Text)
End Sub
