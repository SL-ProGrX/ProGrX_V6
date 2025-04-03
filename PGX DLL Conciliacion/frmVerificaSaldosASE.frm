VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmVerificaSaldosASE 
   Caption         =   "Verificación de Saldos de Créditos"
   ClientHeight    =   6072
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13380
   HelpContextID   =   7007
   Icon            =   "frmVerificaSaldosASE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6072
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3492
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   5652
      _Version        =   1245185
      _ExtentX        =   9970
      _ExtentY        =   6159
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   600
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   5808
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   9720
      TabIndex        =   3
      Top             =   240
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmVerificaSaldosASE.frx":000C
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   7680
      TabIndex        =   4
      Top             =   360
      Width           =   1812
      _Version        =   1245185
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdArchivo 
      Height          =   492
      Left            =   11160
      TabIndex        =   5
      Top             =   240
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Archivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmVerificaSaldosASE.frx":0A2A
   End
   Begin XtremeSuiteControls.ComboBox cboBusca 
      Height          =   312
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   1812
      _Version        =   1245185
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkSaldosInicial 
      Height          =   204
      Left            =   5880
      TabIndex        =   13
      Top             =   720
      Width           =   204
      _Version        =   1245185
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Excluir Operaciones Nuevas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   6120
      TabIndex        =   14
      Top             =   720
      Width           =   2892
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   492
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   5412
      _Version        =   1245185
      _ExtentX        =   9546
      _ExtentY        =   868
      _StockProps     =   79
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Actual"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte de Saldo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar en:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   7680
      TabIndex        =   6
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label lblUltimoCorte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmVerificaSaldosASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFechaInicio As Date, vFechaActual As Date


Private Sub cboBusca_Click()
If cboBusca.Text = "Actual" Then
  cboPeriodos.Visible = False
Else
  cboPeriodos.Visible = True
End If
lsw.ListItems.Clear
lblPeriodo.Visible = cboPeriodos.Visible

End Sub

Private Sub cboBusca_KeyDown(KeyCode As Integer, Shift As Integer)
lsw.ListItems.Clear

If cboBusca.Text = "Actual" Then
  cboPeriodos.Visible = False
Else
  cboPeriodos.Visible = True
End If
End Sub

Private Sub cmdArchivo_Click()
 If cboPeriodos.Text = "" Then Exit Sub
 Call sbListViewExporFileTab(lsw)
End Sub

Private Sub cmdBuscar_Click()


If cboBusca.Text = "Actual" Then
  Call sbBuscaActual
Else
  If cboPeriodos.Text = "" Then
    MsgBox "No se Ha especificado ningún periodo...", vbExclamation
  Else
    Call sbBuscaHistorico
  End If
End If


End Sub


Private Sub sbBuscaHistorico()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngEvaluados As Long, lngDiferencias As Long
Dim lngRegistros As Long, curSaldoFinal As Currency
Dim itmX As ListViewItem, curTotales(5) As Currency

Me.MousePointer = vbHourglass
lsw.ListItems.Clear


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"

strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)


strSQL = "select H.*,S.cedula,S.nombre " _
       & "from ase_per_cerrados H inner join reg_creditos R on H.id_solicitud = R.id_solicitud " _
       & " inner join Socios S on R.cedula = S.cedula " _
       & " inner join Catalogo C on R.codigo = C.codigo " _
       & " where H.saldo_final <> (H.saldo_inicial + H.total_debitos - H.total_creditos) " _
       & " and H.anio = " & rs!Anio & " and H.mes = " & rs!Mes _
       & " and C.poliza = 'N' and C.retencion = 'N'"
       
If chkSaldosInicial.Value = xtpChecked Then
    strSQL = strSQL & " and H.Saldo_Inicial <> 0"
End If
rs.Close

Call OpenRecordSet(rs, strSQL)

With rs
    lngRegistros = .RecordCount
    prg.Max = .RecordCount + 1
    prg.Value = 1
    
    curTotales(1) = 0
    curTotales(2) = 0
    curTotales(3) = 0
    curTotales(4) = 0
    curTotales(5) = 0
    
    Do While Not .EOF
     lngEvaluados = lngEvaluados + 1
     
     'Si la diferencia es mayor de 1
     
     If Abs(!saldo_final - (!saldo_inicial + !total_debitos - !total_creditos)) > 1 Then
         lngDiferencias = lngDiferencias + 1
         
         Set itmX = lsw.ListItems.Add(, , !id_solicitud)
           itmX.SubItems(1) = Trim(!Codigo)
           itmX.SubItems(2) = Trim(!Cedula)
           itmX.SubItems(3) = Trim(!Nombre) & ""
           itmX.SubItems(4) = Format(!saldo_inicial, "Standard")
           itmX.SubItems(5) = Format(!saldo_final, "Standard")
           itmX.SubItems(6) = Format(!total_debitos, "Standard")
           itmX.SubItems(7) = Format(!total_creditos, "Standard")
           itmX.SubItems(8) = Format(!saldo_final - (!saldo_inicial + !total_debitos - !total_creditos), "Standard")
                 
        curTotales(1) = curTotales(1) + !saldo_inicial
        curTotales(2) = curTotales(2) + !saldo_final
        curTotales(3) = curTotales(3) + !total_debitos
        curTotales(4) = curTotales(4) + !total_creditos
        curTotales(5) = curTotales(5) + !saldo_final - (!saldo_inicial + !total_debitos - !total_creditos)
        
     End If
     
     .MoveNext
     lblEstado.Caption = "Procesados  " & Format(lngEvaluados, "###,###,###,##0") _
               & "  De  " & Format(lngRegistros, "###,###,###,##0") & "  -  Porcentaje: " _
               & Round((lngEvaluados / lngRegistros) * 100, 2) & "%" & vbCrLf _
               & "Evaluados " & Format(lngEvaluados, "###,###,###,##0") & "  -  Diferencias " & Format(lngDiferencias, "###,###,###,##0")
    
     prg.Value = prg.Value + 1
     If Right(CStr(prg.Value), 2) = "00" Then DoEvents
    Loop
 .Close
    
       'TOTALES
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = "---------------------"
           itmX.SubItems(5) = "---------------------"
           itmX.SubItems(6) = "---------------------"
           itmX.SubItems(7) = "---------------------"
           itmX.SubItems(8) = "---------------------"
                 
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = Format(curTotales(1), "Standard")
           itmX.SubItems(5) = Format(curTotales(2), "Standard")
           itmX.SubItems(6) = Format(curTotales(3), "Standard")
           itmX.SubItems(7) = Format(curTotales(4), "Standard")
           itmX.SubItems(8) = Format(curTotales(5), "Standard")
                 
End With

Me.MousePointer = vbDefault


End Sub

Private Sub sbBuscaActual()
Dim strSQL As String, rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim curDebitos As Currency, curCreditos As Currency
Dim lngEvaluados As Long, lngDiferencias As Long
Dim lngRegistros As Long, curSaldoFinal As Currency
Dim itmX As ListViewItem, curTotales(5) As Currency


Me.MousePointer = vbHourglass
lsw.ListItems.Clear


lblEstado.Caption = vbCrLf & "****- Actualizado operaciones nuevas -****"


strSQL = "update reg_creditos set saldo_inicial = montoapr" _
       & " where fechaforp between '" & Format(vFechaInicio, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(vFechaActual, "yyyy/mm/dd") & " 23:59:59' and saldo_inicial = 0" _
       & " and estado in('A','C','N')"
Call ConectionExecute(strSQL)


lblEstado.Caption = vbCrLf & "****- Cargando Información, Espere -****"



'Se excluyen las retenciones y polizas (PSD) de la busqueda
strSQL = "Select R.id_solicitud,R.codigo,R.saldo,R.saldo_inicial,S.cedula,S.nombre" _
       & " From Reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join Catalogo C on R.codigo = C.codigo" _
       & " Where Saldo is not null and Saldo <> Saldo_inicial and estado in('A','C')" _
       & " and C.retencion = 'N' and C.Poliza = 'N'" _
       & " and R.montoapr > R.saldo"
      ' & " and fechaforp < '" & Format(dtp1.Value, "yyyy/mm/dd") & "'"
If chkSaldosInicial.Value = xtpChecked Then
    strSQL = strSQL & " and R.Saldo_Inicial <> 0"
End If

Call OpenRecordSet(rs, strSQL)

With rs
    lngRegistros = .RecordCount
    prg.Max = .RecordCount + 1
    prg.Value = 1
    
    curTotales(1) = 0
    curTotales(2) = 0
    curTotales(3) = 0
    curTotales(4) = 0
    curTotales(5) = 0
    
    Do While Not .EOF
     lngEvaluados = lngEvaluados + 1
     
     If !saldo <> !saldo_inicial Then
       curDebitos = 0
       curCreditos = 0
       
       strSQL = "Select * From Creditos_dt where id_solicitud=" & !id_solicitud
       strSQL = strSQL & " And Fechas between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' And '" _
              & Format(vFechaActual, "yyyy/mm/dd") & " 23:59:59'"
       rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         Do While Not rsTmp.EOF
            Select Case rsTmp!tcon
              Case 8
                curDebitos = curDebitos + rsTmp!amortiza
              Case Else
                curCreditos = curCreditos + rsTmp!amortiza
            End Select
            rsTmp.MoveNext
         Loop
       rsTmp.Close
              
       strSQL = "Select * From Morosidad where id_solicitud=" & !id_solicitud _
              & " And Fecult between '" & Format(vFechaInicio, "yyyy/mm/dd") _
              & "' And '" & Format(vFechaActual, "yyyy/mm/dd") & "' and estado <> 'A'"
       rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         Do While Not rsTmp.EOF
            Select Case rsTmp!tcon
              Case 8
                curDebitos = curDebitos + IIf(IsNull(rsTmp!abAmortiza), 0, rsTmp!abAmortiza)
              Case Else
                curCreditos = curCreditos + IIf(IsNull(rsTmp!abAmortiza), 0, rsTmp!abAmortiza)
            End Select
            rsTmp.MoveNext
         Loop
       rsTmp.Close
       
       curSaldoFinal = !saldo_inicial + curDebitos - curCreditos
       If Abs(curSaldoFinal - !saldo) > 1 Then
         
         lngDiferencias = lngDiferencias + 1
         
         Set itmX = lsw.ListItems.Add(, , !id_solicitud)
           itmX.SubItems(1) = Trim(!Codigo)
           itmX.SubItems(2) = Trim(!Cedula)
           itmX.SubItems(3) = Trim(!Nombre) & ""
           itmX.SubItems(4) = Format(!saldo_inicial, "Standard")
           itmX.SubItems(5) = Format(!saldo, "Standard")
           itmX.SubItems(6) = Format(curDebitos, "Standard")
           itmX.SubItems(7) = Format(curCreditos, "Standard")
           itmX.SubItems(8) = Format(!saldo - curSaldoFinal, "Standard")
                 
        curTotales(1) = curTotales(1) + !saldo_inicial
        curTotales(2) = curTotales(2) + !saldo
        curTotales(3) = curTotales(3) + curDebitos
        curTotales(4) = curTotales(4) + curCreditos
        curTotales(5) = curTotales(5) + !saldo - curSaldoFinal
        
       
                 
       End If
     End If
     
     .MoveNext
     lblEstado.Caption = "Procesados  " & Format(lngEvaluados, "###,###,###,##0") _
               & "  De  " & Format(lngRegistros, "###,###,###,##0") & "  -  Porcentaje: " _
               & Round((lngEvaluados / lngRegistros) * 100, 2) & "%" & vbCrLf _
               & "Evaluados " & Format(lngEvaluados, "###,###,###,##0") & "  -  Diferencias " & Format(lngDiferencias, "###,###,###,##0")
    
     prg.Value = prg.Value + 1
     If Right(CStr(prg.Value), 2) = "00" Then DoEvents
    Loop
 .Close
    
       'TOTALES
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = "---------------------"
           itmX.SubItems(5) = "---------------------"
           itmX.SubItems(6) = "---------------------"
           itmX.SubItems(7) = "---------------------"
           itmX.SubItems(8) = "---------------------"
                 
         Set itmX = lsw.ListItems.Add()
           itmX.SubItems(4) = Format(curTotales(1), "Standard")
           itmX.SubItems(5) = Format(curTotales(2), "Standard")
           itmX.SubItems(6) = Format(curTotales(3), "Standard")
           itmX.SubItems(7) = Format(curTotales(4), "Standard")
           itmX.SubItems(8) = Format(curTotales(5), "Standard")
                 
End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
       .Clear
       .Add , , "Operación", 1900
       .Add , , "Código", 1200, vbCenter
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "Saldo Inicial", 2100, vbRightJustify
       .Add , , "Saldo Final", 2100, vbRightJustify
       .Add , , "Débitos", 2100, vbRightJustify
       .Add , , "Créditos", 2100, vbRightJustify
       .Add , , "Diferencia", 2100, vbRightJustify
End With


cboBusca.Clear
cboBusca.AddItem "Actual"
cboBusca.AddItem "Histórico"
cboBusca.Text = "Actual"

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - (lsw.Top + lblEstado.Height + prg.Height + 480)
lblEstado.Top = lsw.Top + lsw.Height + 20
lblEstado.Width = lsw.Width

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub sbInicial()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select valor as Fecha,dbo.MyGetdate() as Actual" _
       & " from crd_parametros where cod_parametro = '16'"

Call OpenRecordSet(rs, strSQL)
vFechaInicio = Format(rs!Fecha, "yyyy/mm/dd")
vFechaActual = rs!actual
lblUltimoCorte.Caption = Format(rs!Fecha, "dd/mmm/yyyy")
lblFecha.Caption = Format(rs!actual, "dd/mmm/yyyy")
rs.Close



strSQL = "select * from ase_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)


cboPeriodos.Clear
Do While Not rs.EOF
 cboPeriodos.AddItem rs!Anio & " - " & fxConvierteMES(rs!Mes)
 cboPeriodos.ItemData(cboPeriodos.ListCount - 1) = CStr(rs!id_per_historico)
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
 rs.MoveFirst
 Call sbCboAsignaDato(cboPeriodos, rs!Anio & " - " & fxConvierteMES(rs!Mes), True, rs!id_per_historico)
End If
rs.Close

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub

