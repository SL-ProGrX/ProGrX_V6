VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_Monitoreo 
   Caption         =   "Monitoreo de Saldos"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   Icon            =   "frmTES_Monitoreo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   8070
      _StockProps     =   77
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Left            =   2400
      Top             =   480
   End
   Begin VB.CheckBox chkDetener 
      Caption         =   "Deterner Monitoreo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
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
   Begin XtremeSuiteControls.ComboBox cboTipoMov 
      Height          =   330
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
      Caption         =   "Tipo de Movimiento"
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
      Left            =   5520
      TabIndex        =   6
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblActualiza 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">>> Actualizando <<<"
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
      Height          =   252
      Left            =   6720
      TabIndex        =   2
      Top             =   624
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Corte"
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
      Height          =   195
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   11772
   End
End
Attribute VB_Name = "frmTES_Monitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDetener_Click()
If chkDetener.Value = vbChecked Then
  chkDetener.BackColor = vbRed
  TimerX.Interval = 0
Else
  chkDetener.BackColor = RGB(70, 111, 178)
  TimerX.Interval = 60000
End If
End Sub

Private Sub Form_Activate()
 vModulo = 9

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, rsTmp As New ADODB.Recordset


vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboTipoMov.AddItem "Según movimientos cargados"
cboTipoMov.AddItem "Según consulta de desembolsos"
cboTipoMov.Text = "Según consulta de desembolsos"


chkDetener.BackColor = RGB(70, 111, 178)

dtpFecha.Value = fxFechaServidor

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Corte", 1200, vbCenter
    .Add , , "Cuenta", 2200
    .Add , , "Descripción", 3500
    .Add , , "Saldo Inicial", 2100, vbRightJustify
    .Add , , "Total Débitos", 2100, vbRightJustify
    .Add , , "Total Crédito", 2100, vbRightJustify
    
    .Add , , "Cheques Pendientes", 2100, vbRightJustify
    .Add , , "Cheques Día", 2100, vbRightJustify
    .Add , , "Transferencia Dia ", 2100, vbRightJustify
    
    .Add , , "Saldo Final", 2100, vbRightJustify
    .Add , , "Saldo Mínimo", 2100, vbRightJustify
    .Add , , "Diferencia Saldo", 2100, vbRightJustify
    
'   .Add , , "Cuenta Conta", 2500, vbCenter
    

End With

lsw.Checkboxes = False
lsw.ListItems.Clear



strSQL = "select H.id_banco,isnull(max(H.idX),0) as IDX" _
       & " from TES_BANCOS_CIERRES H inner join Tes_Bancos B on H.id_banco = B.id_banco" _
       & " Where B.monitoreo = 1 group by H.id_banco"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 strSQL = "select H.id_banco,B.descripcion,B.cta,H.corte,H.saldo_final,B.ctaConta,H.saldo_minimo" _
        & " from TES_BANCOS_CIERRES H inner join Tes_Bancos B on H.id_banco = B.id_banco" _
        & " Where H.idx = " & rs!IdX
 Call OpenRecordSet(rsTmp, strSQL, 0)
 If Not rsTmp.EOF And Not rsTmp.BOF Then
    Set itmX = lsw.ListItems.Add(, , rs!Id_Banco)
        itmX.SubItems(1) = rsTmp!Cta
        itmX.SubItems(2) = rsTmp!DESCRIPCION
        itmX.SubItems(3) = Format(DateAdd("d", 1, rsTmp!Corte), "yyyy/mm/dd")
        itmX.SubItems(4) = Format(rsTmp!saldo_final, "Standard")
        itmX.SubItems(5) = 0
        itmX.SubItems(6) = 0
        itmX.SubItems(7) = Format(rsTmp!saldo_final, "Standard")
        itmX.SubItems(8) = fxCntX_CuentaFormato(True, Trim(rsTmp!ctaConta), 0)
        itmX.SubItems(9) = Format(rsTmp!saldo_minimo, "Standard")
 End If
 rsTmp.Close
 
 rs.MoveNext
Loop
rs.Close

TimerX.Interval = 10

End Sub

Private Sub Form_Resize()
On Error GoTo vError
    imgBanner.Width = Me.Width
    
    lsw.Width = Me.Width - 150
    lsw.Height = Me.Height - (imgBanner.Height + 150)
vError:
End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

TimerX.Interval = 0
lblActualiza.Visible = True
lblActualiza.Refresh

With lsw.ListItems
 For i = 1 To .Count
  
    .Item(i).SubItems(5) = 0
    .Item(i).SubItems(6) = 0
    
    'Emisiones de Documentos
    strSQL = "select D.debehaber as Movimiento,sum(D.monto / D.Tipo_Cambio) as Total" _
           & " from Tes_Transacciones C inner join Tes_Trans_Asiento D on C.nsolicitud = D.nsolicitud" _
           & " where C.fecha_emision between '" & Format(DateAdd("d", -1, .Item(i).SubItems(3)), "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpFecha.Value, "yyyy/mm/dd") _
           & " 23:59:59' and C.estado in('I','T','A') and D.cuenta_contable = '" & fxCntX_CuentaFormato(False, .Item(i).SubItems(8), 0) _
           & "' group by D.debehaber"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     If rs!Movimiento = "D" Then
        .Item(i).SubItems(6) = rs!Total
     Else
        .Item(i).SubItems(5) = rs!Total
     End If
     rs.MoveNext
    Loop
    rs.Close
    
    
    'Anulaciones de Documentos
    strSQL = "select D.debehaber as Movimiento,sum(D.monto/ D.Tipo_Cambio) as Total" _
           & " from Tes_Transacciones C inner join Tes_Trans_Asiento D on C.nsolicitud = D.nsolicitud" _
           & " where C.fecha_anula between '" & Format(DateAdd("d", -1, .Item(i).SubItems(3)), "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpFecha.Value, "yyyy/mm/dd") _
           & " 23:59:59' and C.estado in('A') and D.cuenta_contable = '" & fxCntX_CuentaFormato(False, .Item(i).SubItems(8), 0) _
           & "' group by D.debehaber"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     If rs!Movimiento = "D" Then
        .Item(i).SubItems(5) = CCur(.Item(i).SubItems(5)) + rs!Total
     Else
        .Item(i).SubItems(6) = CCur(.Item(i).SubItems(6)) + rs!Total
     End If
     rs.MoveNext
    Loop
    rs.Close
    
    .Item(i).SubItems(5) = Format(.Item(i).SubItems(5), "Standard")
    .Item(i).SubItems(6) = Format(.Item(i).SubItems(6), "Standard")
    
    .Item(i).SubItems(7) = Format(CCur(.Item(i).SubItems(4)) - CCur(.Item(i).SubItems(5)) + CCur(.Item(i).SubItems(6)), "Standard")
     
    'Verde : Saldo > al mínimo
    If CCur(.Item(i).SubItems(7)) >= CCur(.Item(i).SubItems(9)) Then
       .Item(i).TextBackColor = vbWhite
       .Item(i).Bold = False
    End If
     
     
    'Amarillo : Saldo > 0 y Saldo por Debajo del mínimo
    If CCur(.Item(i).SubItems(7)) < CCur(.Item(i).SubItems(9)) And CCur(.Item(i).SubItems(7)) > 0 Then
       .Item(i).Bold = False
       .Item(i).TextBackColor = RGB(252, 243, 207) 'Amarillo
    End If
    
    'Rojo : Saldo <= 0
    If CCur(.Item(i).SubItems(7)) <= 0 Then
        .Item(i).ForeColor = vbRed
        .Item(i).Bold = True
        .Item(i).TextBackColor = RGB(250, 219, 216) 'Rojo
    End If
     
 Next i
End With

lblActualiza.Visible = False
Me.MousePointer = vbDefault
TimerX.Interval = 60000


End Sub
