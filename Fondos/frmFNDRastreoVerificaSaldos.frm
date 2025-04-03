VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmFNDRastreoVerificaSaldos 
   Caption         =   "Verificación de Saldos del los Fondos"
   ClientHeight    =   4644
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4644
   ScaleWidth      =   11808
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2772
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8652
      _Version        =   1245185
      _ExtentX        =   15261
      _ExtentY        =   4890
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
   Begin XtremeSuiteControls.CheckBox chkDiferencias 
      Height          =   492
      Left            =   7680
      TabIndex        =   11
      Top             =   240
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Diferencias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   4368
      Width           =   11808
      _ExtentX        =   20828
      _ExtentY        =   487
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   4680
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
      Picture         =   "frmFNDRastreoVerificaSaldos.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   3372
      _Version        =   1245185
      _ExtentX        =   5948
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
      Left            =   6120
      TabIndex        =   5
      Top             =   240
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmFNDRastreoVerificaSaldos.frx":0A1E
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   372
      Left            =   9360
      TabIndex        =   6
      Top             =   360
      Width           =   732
      _Version        =   1245185
      _ExtentX        =   1291
      _ExtentY        =   656
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
      Text            =   "30000"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   312
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   3372
      _Version        =   1245185
      _ExtentX        =   5948
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas"
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
      Height          =   312
      Index           =   3
      Left            =   9360
      TabIndex        =   10
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Height          =   312
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
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
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1212
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstado 
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   14412
      _Version        =   1245185
      _ExtentX        =   25421
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmFNDRastreoVerificaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxSaldoFinal(vOperadora As Long, vCodPlan As String, vContrato As Long _
        , vAnio As Long, vMes As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select aportes,rendimientos from fnd_per_cerrados where anio = " & vAnio _
       & " and mes = " & vMes & " and cod_operadora = " & vOperadora & " and cod_plan = '" _
       & vCodPlan & "' and cod_contrato = " & vContrato
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxSaldoFinal = 0
Else
  fxSaldoFinal = rs!aportes + rs!rendimientos
End If
rs.Close

End Function


Private Function fxMovimientos(vOperadora As Long, vPlan As String, vContrato As Long _
        , vFechaInicio As Date, vFechaCorte As Date, Optional vTipo As String = "D")
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select abs(isnull(sum(D.monto),0)) as Monto" _
       & " from fnd_contratos C inner join fnd_contratos_detalle D on C.cod_operadora = D.cod_operadora" _
       & " and C.cod_plan = D.cod_plan and C.cod_contrato = D.cod_contrato" _
       & " where D.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
       & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and D.cod_operadora = " _
       & vOperadora & " and D.cod_plan = '" & vPlan & "' and D.cod_contrato = " & vContrato
If vTipo = "D" Then
   strSQL = strSQL & " and D.Monto < 0"
Else
   strSQL = strSQL & " and D.Monto > 0"
End If
Call OpenRecordSet(rs, strSQL)
fxMovimientos = rs!Monto
rs.Close

End Function


Private Sub sbDetalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotales(5) As Currency
Dim lngLineas As Long, lngRegistros As Long, lngEvaluados As Long
Dim vFechaInicio As Date, vFechaCorte As Date, vPaso As Boolean
Dim lngAnio As Long, iMes As Integer, lngAnioX As Long, iMesX As Integer
Dim curSaldoInicial As Currency, curSaldoFinal As Currency
Dim curDebitos As Currency, curCreditos As Currency

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

lblEstado.Caption = "Cargando datos y configurando periodos de busqueda..."

strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)
    lngAnio = rs!Anio
    iMes = rs!Mes
rs.Close

'Fecha de Inicio y Corte del Periodo a Busca, para los movimientos
vFechaInicio = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
If iMes = 12 Then
   iMesX = 1
   lngAnioX = lngAnio + 1
Else
   iMesX = iMes + 1
   lngAnioX = lngAnio
End If
   
vFechaCorte = CDate(lngAnioX & "/" & Format(iMesX, "00") & "/01")
vFechaCorte = DateAdd("d", -1, vFechaCorte)
   
'Fecha de Proceso Anterior
If iMes = 1 Then
   iMesX = 12
   lngAnioX = lngAnio - 1
Else
   iMesX = iMes - 1
   lngAnioX = lngAnio
End If
   
   
'Inicializa los totales
curTotales(0) = 0 'Saldo Inicial
curTotales(1) = 0 'Debitos
curTotales(2) = 0 'Creditos
curTotales(3) = 0 'Resultado
curTotales(4) = 0 'Saldo Final
curTotales(5) = 0 'Diferencia

lngEvaluados = 0
lngLineas = 1

'Ver si A Futuro se realiza por Codigo de Plan
strSQL = "select TOP " & txtLineas.Text & " C.cod_operadora,C.cod_plan,C.cod_contrato,C.cedula,S.nombre,(D.Aportes+D.rendimientos) as 'Saldo_Final'" _
       & ", isnull(M.Debito,0) as 'DEBITOS', isnull(M.Credito,0) as 'CREDITOS'" _
       & ", isnull(A.Aportes + A.Rendimientos,0) as 'Saldo_Inicial'" _
       & " from fnd_contratos C " _
       & " inner join fnd_per_cerrados D on C.cod_operadora = D.cod_operadora" _
       & " and C.cod_plan = D.cod_plan and C.cod_contrato = D.cod_contrato" _
       & " and D.anio = " & lngAnio & " and D.mes = " & iMes _
       & " left join Socios S on C.cedula = S.cedula" _
       & " left join fnd_per_cerrados A on D.cod_operadora = A.cod_Operadora" _
       & "   and D.cod_Plan = A.cod_Plan and D.cod_Contrato = A.cod_Contrato" _
       & "   and A.anio = " & lngAnioX & " and A.mes = " & iMesX _
       & " left join vFnd_Contratos_Mov_Periodo_Rsm M on D.cod_operadora = M.cod_Operadora" _
       & "   and D.cod_Plan = M.cod_Plan and D.cod_Contrato = M.cod_Contrato" _
       & "   and D.Anio = M.Anio and D.Mes = M.Mes"



If cboPlan.Text <> "TODOS" Then
 strSQL = strSQL & " Where D.cod_plan ='" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
End If

Call OpenRecordSet(rs, strSQL)
prg.Max = rs.RecordCount + 2
prg.Value = 1


Do While Not rs.EOF

 lngEvaluados = lngEvaluados + 1

 curSaldoFinal = rs!saldo_final
 curSaldoInicial = rs!saldo_inicial
 curDebitos = rs!Debitos
 curCreditos = rs!Creditos
 
 If chkDiferencias.Value = vbChecked Then
   If (curSaldoInicial + curCreditos - curDebitos) <> curSaldoFinal Then
      vPaso = True
   Else
      vPaso = False
   End If
 Else
   vPaso = True
 End If
  
  
 If vPaso Then
    curTotales(0) = curTotales(0) + curSaldoInicial    'Saldo Inicial
    curTotales(1) = curTotales(1) + curDebitos 'Debitos
    curTotales(2) = curTotales(2) + curCreditos 'Creditos
    curTotales(3) = curTotales(3) + (curSaldoInicial + curCreditos - curDebitos) 'Resultado
    curTotales(4) = curTotales(4) + curSaldoFinal 'Saldo Final
    curTotales(5) = curSaldoFinal - (curSaldoInicial + curCreditos - curDebitos) 'Diferencia
   
   Set itmX = lsw.ListItems.Add(, , rs!Cod_Operadora)
       itmX.SubItems(1) = rs!cod_Plan
       itmX.SubItems(2) = rs!COD_CONTRATO
       itmX.SubItems(3) = Format(curSaldoInicial, "Standard")
       itmX.SubItems(4) = Format(curDebitos, "Standard")
       itmX.SubItems(5) = Format(curCreditos, "Standard")
       itmX.SubItems(6) = Format((curSaldoInicial + curCreditos - curDebitos), "Standard")
       itmX.SubItems(7) = Format(curSaldoFinal, "Standard")
       itmX.SubItems(8) = Format(CCur(itmX.SubItems(7)) - CCur(itmX.SubItems(6)), "Standard")
       itmX.SubItems(9) = rs!Cedula
       itmX.SubItems(10) = rs!Nombre
       
       If CCur(itmX.SubItems(8)) <> 0 Then
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216)
       End If
       
       lngLineas = lngLineas + 1
 End If
 
 prg.Value = prg.Value + 1
 
 lblEstado.Caption = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(prg.Max, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(vFechaInicio, "dd/mm/yyyy") & " al " & Format(vFechaCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents
 rs.MoveNext
Loop
rs.Close


'TOTALES
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(3) = "_______________"
    itmX.SubItems(4) = "_______________"
    itmX.SubItems(5) = "_______________"
    itmX.SubItems(6) = "_______________"
    itmX.SubItems(7) = "_______________"
    itmX.SubItems(8) = "_______________"
          
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(3) = Format(curTotales(0), "Standard")
    itmX.SubItems(4) = Format(curTotales(1), "Standard")
    itmX.SubItems(5) = Format(curTotales(2), "Standard")
    itmX.SubItems(6) = Format(curTotales(3), "Standard")
    itmX.SubItems(7) = Format(curTotales(4), "Standard")
    itmX.SubItems(8) = Format(curTotales(5), "Standard")


Me.MousePointer = vbDefault

lblEstado.Caption = "Consulta Finalizada..."

End Sub



Private Sub sbArchivo()
Dim strCadena As String, fn, lng As Long, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngLineas As Long, lngRegistros As Long, lngEvaluados As Long
Dim vFechaInicio As Date, vFechaCorte As Date, vPaso As Boolean
Dim lngAnio As Long, iMes As Integer, lngAnioX As Long, iMesX As Integer
Dim curSaldoInicial As Currency, curSaldoFinal As Currency
Dim curDebitos As Currency, curCreditos As Currency

Dim archivo As String

Me.MousePointer = vbHourglass


fn = FreeFile

archivo = SIFGlobal.DirectorioDeResultados & "\FNDVerificaSaldos.txt"

Open archivo For Output As #fn  ' Crea Archivo.

'Imprimir Encabezados
Print #fn, "Verificación de Saldos de los Fondos" & vbCrLf _
            & "Periodo : " & cboPeriodos.Text

strCadena = ""
For lng = 1 To lsw.ColumnHeaders.Count
  strCadena = strCadena & UCase(lsw.ColumnHeaders.Item(lng).Text) & vbTab
Next lng
Print #fn, strCadena

lblEstado.Caption = "Cargando datos y configurando periodos de busqueda..."

strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)
    lngAnio = rs!Anio
    iMes = rs!Mes
rs.Close

'Fecha de Inicio y Corte del Periodo a Busca, para los movimientos
vFechaInicio = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
If iMes = 12 Then
   iMesX = 1
   lngAnioX = lngAnio + 1
Else
   iMesX = iMes + 1
   lngAnioX = lngAnio
End If
   
vFechaCorte = CDate(lngAnioX & "/" & Format(iMesX, "00") & "/01")
vFechaCorte = DateAdd("d", -1, vFechaCorte)
   
'Fecha de Proceso Anterior
If iMes = 1 Then
   iMesX = 12
   lngAnioX = lngAnio - 1
Else
   iMesX = iMes - 1
   lngAnioX = lngAnio
End If
   

lngEvaluados = 0
lngLineas = 1

'Ver si A Futuro se realiza por Codigo de Plan
strSQL = "select C.cod_operadora,C.cod_plan,C.cod_contrato,C.cedula,S.nombre,(D.Aportes+D.rendimientos) as 'Saldo_Final'" _
       & ", isnull(M.Debito,0) as 'DEBITOS', isnull(M.Credito,0) as 'CREDITOS'" _
       & ", isnull(A.Aportes + A.Rendimientos,0) as 'Saldo_Inicial'" _
       & " from fnd_contratos C " _
       & " inner join fnd_per_cerrados D on C.cod_operadora = D.cod_operadora" _
       & " and C.cod_plan = D.cod_plan and C.cod_contrato = D.cod_contrato" _
       & " and D.anio = " & lngAnio & " and D.mes = " & iMes _
       & " left join Socios S on C.cedula = S.cedula" _
       & " left join fnd_per_cerrados A on D.cod_operadora = A.cod_Operadora" _
       & "   and D.cod_Plan = A.cod_Plan and D.cod_Contrato = A.cod_Contrato" _
       & "   and A.anio = " & lngAnioX & " and A.mes = " & iMesX _
       & " left join vFnd_Contratos_Mov_Periodo_Rsm M on D.cod_operadora = M.cod_Operadora" _
       & "   and D.cod_Plan = M.cod_Plan and D.cod_Contrato = M.cod_Contrato" _
       & "   and D.Anio = M.Anio and D.Mes = M.Mes"



If cboPlan.Text <> "TODOS" Then
 strSQL = strSQL & " Where D.cod_plan ='" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
End If

Call OpenRecordSet(rs, strSQL)
prg.Max = rs.RecordCount + 2
prg.Value = 1


Do While Not rs.EOF

 lngEvaluados = lngEvaluados + 1

 curSaldoFinal = rs!saldo_final
 curSaldoInicial = rs!saldo_inicial
 curDebitos = rs!Debitos
 curCreditos = rs!Creditos
 
 
 If chkDiferencias.Value = vbChecked Then
   If (curSaldoInicial + curCreditos - curDebitos) <> curSaldoFinal Then
      vPaso = True
   Else
      vPaso = False
   End If
 Else
   vPaso = True
 End If
  
  
 If vPaso Then
   strCadena = rs!Cod_Operadora & vbTab & rs!cod_Plan & vbTab & rs!COD_CONTRATO _
             & vbTab & Format(curSaldoInicial, "Standard") & vbTab & Format(curDebitos, "Standard") _
             & vbTab & Format(curCreditos, "Standard") & vbTab & Format((curSaldoInicial + curCreditos - curDebitos), "Standard") _
             & vbTab & Format(curSaldoFinal, "Standard") & vbTab & Format(curSaldoFinal - (curSaldoInicial + curCreditos - curDebitos), "Standard") _
             & vbTab & rs!Cedula & vbTab & rs!Nombre
   Print #fn, strCadena
 End If
 
 prg.Value = prg.Value + 1
 
 lblEstado.Caption = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(prg.Max, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(vFechaInicio, "dd/mm/yyyy") & " al " & Format(vFechaCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents
 rs.MoveNext
Loop
rs.Close


'Cierra Archivo
Close #fn

Me.MousePointer = vbDefault

MsgBox "Se Creo Archivo de Texto con Tabulaciones: " & archivo, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 On Error Resume Next
 Close #fn

End Sub

Private Sub cmdArchivo_Click()

If cboPeriodos.ListCount <= 0 Then Exit Sub

Call sbArchivo

End Sub


Private Sub cmdBuscar_Click()

If cboPeriodos.ListCount <= 0 Then Exit Sub

Call sbDetalle

End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String


vModulo = 18 'Fondo de Inversion

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

strSQL = "select cod_plan as 'IdX', Descripcion as  'ItmX' from fnd_planes"
Call sbCbo_Llena_New(cboPlan, strSQL, True, True)

strSQL = "select id_per_historico as 'IdX', dbo.fxSys_FechaAnioMesToDatetime(anio,mes) as 'ItmX'" _
       & " from fnd_per_historico" _
       & " order by anio desc,mes desc"
Call sbCbo_Llena_New(cboPeriodos, strSQL, False, True)

With lsw.ColumnHeaders
    .Clear
    .Add , , "Operadora Id", 1100
    .Add , , "Plan Id", 1100, vbCenter
    .Add , , "Contrato Id", 1100
    .Add , , "Saldo Inicial", 2400, vbRightJustify
    .Add , , "Débitos", 2400, vbRightJustify
    .Add , , "Créditos", 2400, vbRightJustify
    .Add , , "SF Calculado", 2400, vbRightJustify
    .Add , , "Saldo Final", 2400, vbRightJustify
    .Add , , "Diferencia", 2400, vbRightJustify
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 4800
    
End With



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 100
lsw.Height = Me.Height - (500 + lblEstado.Height + prg.Height + lsw.top)
lblEstado.top = lsw.top + lsw.Height + 20
lblEstado.Width = lsw.Width
End Sub



