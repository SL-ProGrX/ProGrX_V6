VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmActivos_CierrePeriodo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre del Periodo"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7092
      _Version        =   1441792
      _ExtentX        =   12509
      _ExtentY        =   2773
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdCierre 
         Height          =   1092
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Appearance      =   16
         Picture         =   "frmActivo_CierrePeriodo.frx":0000
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmActivo_CierrePeriodo.frx":09C3
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1092
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   4692
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   2952
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpPeriodo 
      Height          =   312
      Left            =   3600
      TabIndex        =   6
      Top             =   480
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
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
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del Periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   5040
      TabIndex        =   2
      Top             =   480
      Width           =   3612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo a Cerrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmActivos_CierrePeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim mDepAcumulada As Currency, mDepMensual As Currency
'
'
'
'Private Sub sbLineaRecta(vBase As Currency, vVidaUtil As Integer _
'                 , vFechaAd As Date, vFechaCa As Date, DepAcum As Currency _
'                 , vUltPer As Long, Optional vTipoVidaUtil As String = "A")
'
'Dim vPerCorte As Long, vAnio As Integer, vMes As Byte
'Dim vDias As Long, vPaso As Boolean, vAjuste As Currency
'
''Utiliza Varibles de Modulo
'
'vPerCorte = Year(vFechaCa) & Format(Month(vFechaCa), "00")
'mDepAcumulada = DepAcum
'mDepMensual = 0
'
''Antes de la Asignacion, siguiente
'If vPerCorte = vUltPer Then Exit Sub
'
''Convierte la Vida Util a Meses
'If UCase(vTipoVidaUtil) = "A" Then
'  vVidaUtil = vVidaUtil * 12
'End If
'
'If (vBase - DepAcum) <= 0 Then Exit Sub
'
'Do While vUltPer < vPerCorte
'   'Inicializa variable de control
'   vPaso = False
'
'   'Establece el periodo inicial / Mes de Inicio
'   If vUltPer = 0 Then
'        vUltPer = Year(vFechaAd) & Format(Month(vFechaAd), "00")
'        vDias = 30 - Day(vFechaAd)
'        If vDias <= 0 Then
'           vDias = 1
'        End If
'        mDepMensual = vDias * (vBase / (vVidaUtil * 30))
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'   'Esta en el Mes de Corte
'   vAnio = Year(DateAdd("m", vVidaUtil, vFechaAd))
'   vMes = Month(DateAdd("m", vVidaUtil, vFechaAd))
'   If CLng(vAnio & Format(vMes, "00")) = vPerCorte And Not vPaso Then
'        vDias = Day(vFechaAd)
'        mDepMensual = vDias * (vBase / (vVidaUtil * 30))
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'   'Otro Mes Entero
'   If Not vPaso Then
'        vDias = 30
'        mDepMensual = vDias * (vBase / (vVidaUtil * 30))
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'   'Este procedimiento es para corregir datos migrados con malos calculos
'   'de la depreciacion, o cualquier inconsistencia en esta
'   If mDepAcumulada > vBase Then
'      If mDepMensual < (mDepAcumulada - vBase) Then
'         mDepMensual = 0
'         mDepAcumulada = vBase
'      Else
'        'Ajusta el Ultimo Mes
'        vAjuste = mDepMensual - (mDepAcumulada - vBase)
'        mDepMensual = mDepMensual - vAjuste
'        mDepAcumulada = mDepAcumulada - vAjuste
'      End If
'      'Sale del ciclo (Para Ambos Casos)
'      Exit Do
'   End If
'
'   'Periodo Siguiente
'   vAnio = Mid(CStr(vUltPer), 1, 4)
'   vMes = Mid(CStr(vUltPer), 5, 2)
'   If vMes = 12 Then
'      vMes = 1
'      vAnio = vAnio + 1
'   Else
'      vMes = vMes + 1
'   End If
'   vUltPer = vAnio & Format(vMes, "00")
'Loop
'
'mDepAcumulada = Round(mDepAcumulada, 2)
'mDepMensual = Round(mDepMensual, 2)
'
'End Sub
'
'Private Sub sbSumaDigitos(vBase As Currency, vVidaUtil As Integer _
'                 , vFechaAd As Date, vFechaCa As Date, DepAcum As Currency _
'                 , vUltPer As Long, Optional vTipoVidaUtil As String = "A")
'
'Dim vPerCorte As Long, vAnio As Integer, vMes As Byte
'Dim vDias As Long, vPaso As Boolean, vAjuste As Currency
'Dim vDenominador As Long, vNumerador As Long, lngMeses As Long
'
''Utiliza Varibles de Modulo
'
'vPerCorte = Year(vFechaCa) & Format(Month(vFechaCa), "00")
'mDepAcumulada = DepAcum
'mDepMensual = 0
'
''Antes de la Asignacion, siguiente
'If vPerCorte = vUltPer Then Exit Sub
'
''Convierte la Vida Util a Meses
'If UCase(vTipoVidaUtil) = "A" Then
'  vVidaUtil = vVidaUtil * 12
'End If
'
''Denominador para calculo no es Anual si no mensual (OJO con todos los calculos)
'' fx(Denominador) = (n * (n+1)) / 2
'' fx(Numerador) = (n + 1) - Mes
'vDenominador = (vVidaUtil * (vVidaUtil + 1)) / 2
'
'If (vBase - DepAcum) <= 0 Then Exit Sub
'
''Identifica por cual es el Meses Inicial
'lngMeses = DateDiff("m", vFechaAd, vFechaCa) + 1
'
'Do While vUltPer < vPerCorte
'
'   'Establece el periodo inicial / Mes de Inicio
'   If vUltPer = 0 Then
'        vUltPer = Year(vFechaAd) & Format(Month(vFechaAd), "00")
'        lngMeses = 1
'   End If
'
'  'No deberia pasar,pero por si ocurre
'  If lngMeses > vVidaUtil Then lngMeses = vVidaUtil
'
'  vNumerador = (vVidaUtil + 1) - lngMeses
'
'  Select Case lngMeses
'     Case 1         'Primer Mes
'        'Mes de Inicio
'        vDias = 30 - Day(vFechaAd)
'        If vDias <= 0 Then
'           vDias = 1
'        End If
'        mDepMensual = (((vBase * vNumerador) / vDenominador) / 30) * vDias
'        mDepAcumulada = mDepAcumulada + mDepMensual
'
'     Case vVidaUtil 'Ultimo Mes
'        vDias = Day(vFechaAd)
'        mDepMensual = (((vBase * vNumerador) / vDenominador) / 30) * vDias
'        mDepAcumulada = mDepAcumulada + mDepMensual
'
'     Case Else      'Otros Meses
'        mDepMensual = ((vBase * vNumerador) / vDenominador)
'        mDepAcumulada = mDepAcumulada + mDepMensual
'
'  End Select
'
'  lngMeses = lngMeses + 1
'
'   'Este procedimiento es para corregir datos migrados con malos calculos
'   'de la depreciacion, o cualquier inconsistencia en esta
'   If mDepAcumulada > vBase Then
'      If mDepMensual < (mDepAcumulada - vBase) Then
'         mDepMensual = 0
'         mDepAcumulada = vBase
'      Else
'        'Ajusta el Ultimo Mes
'        vAjuste = mDepMensual - (mDepAcumulada - vBase)
'        mDepMensual = mDepMensual - vAjuste
'        mDepAcumulada = mDepAcumulada - vAjuste
'      End If
'      'Sale del ciclo (Para Ambos Casos)
'      Exit Do
'   End If
'
'   'Periodo Siguiente
'   vAnio = Mid(CStr(vUltPer), 1, 4)
'   vMes = Mid(CStr(vUltPer), 5, 2)
'   If vMes = 12 Then
'      vMes = 1
'      vAnio = vAnio + 1
'   Else
'      vMes = vMes + 1
'   End If
'   vUltPer = vAnio & Format(vMes, "00")
'Loop
'
'mDepAcumulada = Round(mDepAcumulada, 2)
'mDepMensual = Round(mDepMensual, 2)
'
'End Sub
'
'
'Private Sub sbUnidadesProducidas(vBase As Currency, vVidaUtil As Integer _
'                 , vFechaAd As Date, vFechaCa As Date, DepAcum As Currency _
'                 , vUltPer As Long, UdEstimadas As Currency, UdAnual As Currency)
'
'Dim vPerCorte As Long, vAnio As Integer, vMes As Byte
'Dim vDias As Long, vPaso As Boolean, vAjuste As Currency
'Dim dbDepDiaria As Double
'
''Utiliza Varibles de Modulo
'
'vPerCorte = Year(vFechaCa) & Format(Month(vFechaCa), "00")
'mDepAcumulada = DepAcum
'mDepMensual = 0
'
''Antes de la Asignacion, siguiente
'If vPerCorte = vUltPer Then Exit Sub
'
'dbDepDiaria = ((vBase * UdAnual) / UdEstimadas) / 360
'
'If (vBase - DepAcum) <= 0 Then Exit Sub
'
'Do While vUltPer < vPerCorte
'   'Inicializa variable de control
'   vPaso = False
'
'   'Mes de Inicio
'   If vUltPer = 0 Then
'        vUltPer = Year(vFechaAd) & Format(Month(vFechaAd), "00")
'        vDias = 30 - Day(vFechaAd)
'        If vDias <= 0 Then
'           vDias = 1
'        End If
'        mDepMensual = vDias * dbDepDiaria
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'
'   'Esta en el Mes de Corte
'   vAnio = Year(DateAdd("m", vVidaUtil, vFechaAd))
'   vMes = Month(DateAdd("m", vVidaUtil, vFechaAd))
'   If CLng(vAnio & Format(vMes, "00")) = vPerCorte And Not vPaso Then
'        vDias = Day(vFechaAd)
'        mDepMensual = vDias * mDepMensual
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'   'Otro Mes Entero
'   If Not vPaso Then
'        vDias = 30
'        mDepMensual = vDias * mDepMensual
'        mDepAcumulada = mDepAcumulada + mDepMensual
'        vPaso = True
'   End If
'
'   'Este procedimiento es para corregir datos migrados con malos calculos
'   'de la depreciacion, o cualquier inconsistencia en esta
'   If mDepAcumulada > vBase Then
'      If mDepMensual < (mDepAcumulada - vBase) Then
'         mDepMensual = 0
'         mDepAcumulada = vBase
'      Else
'        'Ajusta el Ultimo Mes
'        vAjuste = mDepMensual - (mDepAcumulada - vBase)
'        mDepMensual = mDepMensual - vAjuste
'        mDepAcumulada = mDepAcumulada - vAjuste
'      End If
'      'Sale del ciclo (Para Ambos Casos)
'      Exit Do
'   End If
'
'   'Periodo Siguiente
'   vAnio = Mid(CStr(vUltPer), 1, 4)
'   vMes = Mid(CStr(vUltPer), 5, 2)
'   If vMes = 12 Then
'      vMes = 1
'      vAnio = vAnio + 1
'   Else
'      vMes = vMes + 1
'   End If
'   vUltPer = vAnio & Format(vMes, "00")
'Loop
'
'mDepAcumulada = Round(mDepAcumulada, 2)
'mDepMensual = Round(mDepMensual, 2)
'
'End Sub
'
'
'Private Sub sbDepreciacionPreliminar(vAnio As Long, vMes As Integer)
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim vFecha As Date, i As Integer
'
'vFecha = CDate(vAnio & "/" & Format(vMes, "00") & "/01")
'vFecha = DateAdd("m", 1, vFecha)
'vFecha = DateAdd("d", -1, vFecha)
'
''Preguntar si existe el periodo, de lo contrario crearlo
'strSQL = "select estado from Activos_periodos where anio = " & vAnio _
'       & " and mes = " & vMes
'Call OpenRecordSet(rs, strSQL, 0)
'If rs.EOF And rs.BOF Then
'  strSQL = "insert Activos_periodos(anio,mes,estado,traslado,asientos) values(" & vAnio _
'         & "," & vMes & ",'P','P','P')"
'  Call ConectionExecute(strSQL)
'End If
'rs.Close
'
''Borrar Informacion Anterior
'strSQL = "delete Activos_Cierres_H where anio = " & vAnio & " and mes = " & vMes
'Call ConectionExecute(strSQL)
'
'strSQL = "delete Activos_cierres_adiciones_H where anio = " & vAnio & " and mes = " & vMes
'Call ConectionExecute(strSQL)
'
'
''1. CALCULA DEPRECIACION A LOS ACTIVOS
'strSQL = "select Num_Placa,Valor_Historico,Valor_desecho,fecha_adquisicion,Vida_Util" _
'       & ",Met_depreciacion,Vida_Util_En,depreciacion_periodo,depreciacion_mes" _
'       & ",depreciacion_acum,ud_produccion,ud_anio" _
'       & " From Activos_Principal where estado <> 'R'" _
'       & " and fecha_adquisicion <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
'Call OpenRecordSet(rs, strSQL, 0)
'PrgBar.Max = rs.RecordCount + 1
'PrgBar.Value = 1
'Do While Not rs.EOF
' Select Case rs!met_depreciacion
'  Case "N"
'    mDepAcumulada = 0
'    mDepMensual = 0
'
'  Case "L" 'Linea Recta
'    Call sbLineaRecta((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "S" 'Suma de Digitos
'    Call sbSumaDigitos((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "U" 'Unidades Producidas
'    Call sbUnidadesProducidas((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!ud_produccion, rs!ud_anio)
'
'  Case "D" 'Doblemente Decreciente
'    mDepAcumulada = 0
'    mDepMensual = 0
' End Select
'
' 'Inserta Resultados
' strSQL = "insert Activos_Cierres_H(anio,mes,num_placa,valor_libros,valor_desecho,depreciacion_ac" _
'        & ",depreciacion_mes) values(" & vAnio & "," & vMes & ",'" & rs!num_placa _
'        & "'," & rs!valor_historico & "," & rs!valor_desecho & "," & mDepAcumulada _
'        & "," & mDepMensual & ")"
' Call ConectionExecute(strSQL)
'
' PrgBar.Value = PrgBar.Value + 1
' rs.MoveNext
'Loop
'rs.Close
'PrgBar.Value = 1
'
'
''2. CALCULA DEPRECIACION A LAS ADICIONES Y MEJORAS
'strSQL = "select M.Id_AddRet,A.Num_Placa,M.monto,M.fecha,Meses_Calculo as Vida_Util" _
'       & ",A.Met_depreciacion,'M' as Vida_Util_En,M.depreciacion_periodo,M.depreciacion_mes" _
'       & ",M.depreciacion_acum,A.ud_produccion,A.ud_anio" _
'       & " From Activos_Principal A inner join Activos_retiro_adicion M on A.num_placa = M.num_placa" _
'       & " and M.tipo <> 'R'" _
'       & " where A.estado <> 'R' and M.fecha <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
'Call OpenRecordSet(rs, strSQL, 0)
'PrgBar.Max = rs.RecordCount + 1
'PrgBar.Value = 1
'Do While Not rs.EOF
' Select Case rs!met_depreciacion
'  Case "N"
'    mDepAcumulada = 0
'    mDepMensual = 0
'
'  Case "L" 'Linea Recta
'    Call sbLineaRecta(rs!monto, rs!vida_util, rs!fecha _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "S" 'Suma de Digitos
'    Call sbSumaDigitos(rs!monto, rs!vida_util, rs!fecha _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "U" 'Unidades Producidas
'    Call sbUnidadesProducidas((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!ud_produccion, rs!ud_anio)
'
'  Case "D" 'Doblemente Decreciente
'    mDepAcumulada = 0
'    mDepMensual = 0
'
' End Select
'
' 'Inserta Resultados
' strSQL = "insert Activos_cierres_adiciones_H(anio,mes,num_placa,id,valor_libros,valor_desecho,depreciacion_ac" _
'        & ",depreciacion_mes) values(" & vAnio & "," & vMes & ",'" & rs!num_placa _
'        & "'," & rs!id_AddRet & "," & rs!monto & ",0," & mDepAcumulada _
'        & "," & mDepMensual & ")"
' Call ConectionExecute(strSQL)
'
' PrgBar.Value = PrgBar.Value + 1
' rs.MoveNext
'Loop
'rs.Close
'PrgBar.Value = 1
'
'End Sub
'
'Private Sub sbDepreciacionCierre(vAnio As Long, vMes As Integer)
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim vFecha As Date, i As Integer
'
'vFecha = CDate(vAnio & "/" & Format(vMes, "00") & "/01")
'vFecha = DateAdd("m", 1, vFecha)
'vFecha = DateAdd("d", -1, vFecha)
'
''Preguntar si existe el periodo, de lo contrario crearlo
'strSQL = "select estado from Activos_periodos where anio = " & vAnio _
'       & " and mes = " & vMes
'Call OpenRecordSet(rs, strSQL, 0)
'If rs.EOF And rs.BOF Then
'  strSQL = "insert Activos_periodos(anio,mes,estado,traslado,asientos) values(" & vAnio _
'         & "," & vMes & ",'C','P','P')"
'  Call ConectionExecute(strSQL)
'End If
'rs.Close
'
''Borrar Informacion Anterior
'strSQL = "delete Activos_cierres where anio = " & vAnio & " and mes = " & vMes
'Call ConectionExecute(strSQL)
'
'strSQL = "delete Activos_cierres_adiciones where anio = " & vAnio & " and mes = " & vMes
'Call ConectionExecute(strSQL)
'
'
''1. CALCULA DEPRECIACION A LOS ACTIVOS
'strSQL = "select Num_Placa,Valor_Historico,Valor_desecho,fecha_adquisicion,Vida_Util" _
'       & ",Met_depreciacion,Vida_Util_En,depreciacion_periodo,depreciacion_mes" _
'       & ",depreciacion_acum,ud_produccion,ud_anio" _
'       & " From Activos_Principal where estado <> 'R'" _
'       & " and fecha_adquisicion <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
'Call OpenRecordSet(rs, strSQL, 0)
'PrgBar.Max = rs.RecordCount + 1
'PrgBar.Value = 1
'Do While Not rs.EOF
' Select Case rs!met_depreciacion
'  Case "N"
'    mDepAcumulada = 0
'    mDepMensual = 0
'
'  Case "L" 'Linea Recta
'    Call sbLineaRecta((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "S" 'Suma de Digitos
'    Call sbSumaDigitos((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "U" 'Unidades Producidas
'    Call sbUnidadesProducidas((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!ud_produccion, rs!ud_anio)
'
'  Case "D" 'Doblemente Decreciente
'    mDepAcumulada = 0
'    mDepMensual = 0
' End Select
'
' 'Inserta Resultados
' strSQL = "insert Activos_cierres(anio,mes,num_placa,valor_libros,valor_desecho,depreciacion_ac" _
'        & ",depreciacion_mes) values(" & vAnio & "," & vMes & ",'" & rs!num_placa _
'        & "'," & rs!valor_historico & "," & rs!valor_desecho & "," & mDepAcumulada _
'        & "," & mDepMensual & ")"
' Call ConectionExecute(strSQL)
'
' 'Actualiza Dato de la Depreciacion en Activo
' strSQL = "update Activos_Principal set depreciacion_periodo = " & vAnio & Format(vMes, "00") _
'        & ",depreciacion_Acum = " & mDepAcumulada & ",depreciacion_mes = " & mDepMensual _
'        & " where num_placa = '" & rs!num_placa & "'"
' Call ConectionExecute(strSQL)
'
'
' PrgBar.Value = PrgBar.Value + 1
' rs.MoveNext
'Loop
'rs.Close
'PrgBar.Value = 1
'
'
''2. CALCULA DEPRECIACION A LAS ADICIONES Y MEJORAS
'strSQL = "select M.Id_AddRet,A.Num_Placa,M.monto,M.fecha,Meses_Calculo as Vida_Util" _
'       & ",A.Met_depreciacion,'M' as Vida_Util_En,M.depreciacion_periodo,M.depreciacion_mes" _
'       & ",M.depreciacion_acum,A.ud_produccion,A.ud_anio" _
'       & " From Activos_Principal A inner join Activos_retiro_adicion M on A.num_placa = M.num_placa" _
'       & " and M.tipo <> 'R'" _
'       & " where A.estado <> 'R' and M.fecha <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
'Call OpenRecordSet(rs, strSQL, 0)
'PrgBar.Max = rs.RecordCount + 1
'PrgBar.Value = 1
'Do While Not rs.EOF
' Select Case rs!met_depreciacion
'  Case "N"
'    mDepAcumulada = 0
'    mDepMensual = 0
'
'  Case "L" 'Linea Recta
'    Call sbLineaRecta(rs!monto, rs!vida_util, rs!fecha _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "S" 'Suma de Digitos
'    Call sbSumaDigitos(rs!monto, rs!vida_util, rs!fecha _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!vida_util_en)
'
'  Case "U" 'Unidades Producidas
'    Call sbUnidadesProducidas((rs!valor_historico - rs!valor_desecho), rs!vida_util, rs!fecha_adquisicion _
'                     , vFecha, rs!depreciacion_acum, rs!depreciacion_periodo, rs!ud_produccion, rs!ud_anio)
'
'  Case "D" 'Doblemente Decreciente
'    mDepAcumulada = 0
'    mDepMensual = 0
'
' End Select
'
' 'Inserta Resultados
' strSQL = "insert Activos_cierres_adiciones(anio,mes,num_placa,id_AddRet,valor_libros,valor_desecho,depreciacion_ac" _
'        & ",depreciacion_mes) values(" & vAnio & "," & vMes & ",'" & rs!num_placa _
'        & "'," & rs!id_AddRet & "," & rs!monto & ",0," & mDepAcumulada _
'        & "," & mDepMensual & ")"
' Call ConectionExecute(strSQL)
'
' 'Actualiza Dato de la Depreciacion Mejoras
' strSQL = "update Activos_retiro_adicion set depreciacion_periodo = " & vAnio & Format(vMes, "00") _
'        & ",depreciacion_Acum = " & mDepAcumulada & ",depreciacion_mes = " & mDepMensual _
'        & " where num_placa = '" & rs!num_placa & "' and id_AddRet = " & rs!id_AddRet
' Call ConectionExecute(strSQL)
'
'
' PrgBar.Value = PrgBar.Value + 1
' rs.MoveNext
'Loop
'rs.Close
'PrgBar.Value = 1
'
'
''Generar Asientos de Depreciacion AQUI. ***************************************
'
'
'
'End Sub
'
'Private Sub sbAsientoMov()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim vNumAsiento As String, vDescripcion As String, vAnotacion As String
'Dim vDetalle As String, vTipoActivo As String
'Dim i As Integer, rsTmp As New ADODB.Recordset
'Dim vFecha As Date
'
''La fecha tiene que ser el ultimo dia del mes
'vFecha = DateAdd("m", 1, CDate(Year(dtpPeriodo.Value) & "/" & Format(Month(dtpPeriodo.Value), "00") & "/01"))
'vFecha = DateAdd("d", -1, vFecha)
'
'vDescripcion = "ASIENTO DE DEPRECIACION DEL PERIODO " & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00")
'
''Asiento de Depreciacion Activos
'strSQL = "SELECT T.tipo_activo,T.descripcion,T.cod_cuenta_gastos as CG,T.cod_cuenta_DepAcum as CD" _
'       & ",T.Asiento_Genera,coalesce(sum(C.depreciacion_mes),0) as Depreciacion" _
'       & " from Activos_Auxiliar C inner join Activos_Principal A on C.num_placa = A.num_placa" _
'       & " inner join Activos_tipo_activo T on A.tipo_activo = T.tipo_Activo" _
'       & " where C.anio = " & Year(dtpPeriodo.Value) & " and C.mes = " & Month(dtpPeriodo.Value) _
'       & " group by T.tipo_activo,T.descripcion,T.cod_cuenta_gastos,T.cod_cuenta_DepAcum,T.Asiento_Genera"
'Call OpenRecordSet(rs, strSQL, 0)
'Do While Not rs.EOF
'  vNumAsiento = "AF-C{" & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") & "}T" & Trim(rs!tipo_activo)
'  vAnotacion = "ASIENTO DEPRECIACION RESUMEN X TIPO ACTIVOS"
'  vDetalle = "DEPRECIACION DE " & rs!Descripcion
'  If rs!depreciacion > 0 Then
'    Call sbGAsientoMaestro(rs!Asiento_Genera, vNumAsiento, vFecha, vDescripcion, vAnotacion, "C")
'    Call sbGAsientoDetalle(1, rs!Asiento_Genera, vNumAsiento, rs!CG, rs!depreciacion, "D", "PER." & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") _
'                    , vDetalle, "CIERRE")
'    Call sbGAsientoDetalle(2, rs!Asiento_Genera, vNumAsiento, rs!CD, rs!depreciacion, "H", "PER." & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") _
'                    , vDetalle, "CIERRE")
'  End If
'  rs.MoveNext
'Loop
'rs.Close
'
'
''Asiento de Depreciacion mejoras y revaluaciones
'strSQL = "SELECT J.cod_justificacion,J.descripcion,J.cod_cuenta_03 as CG,J.cod_cuenta_02 as CD" _
'       & ",coalesce(sum(C.depreciacion_mes),0) as Depreciacion" _
'       & " from Activos_Auxiliar_adiciones C inner join Activos_Principal A on C.num_placa = A.num_placa" _
'       & " inner join Activos_retiro_adicion X on C.num_placa = X.num_placa and C.id_AddRet = X.id_AddRet" _
'       & " inner join Activos_justificaciones J on X.cod_justificacion = J.cod_justificacion" _
'       & " where C.anio = " & Year(dtpPeriodo.Value) & " and C.mes = " & Month(dtpPeriodo.Value) _
'       & " group by J.cod_justificacion,J.descripcion,J.cod_cuenta_03,J.cod_cuenta_02"
'Call OpenRecordSet(rs, strSQL, 0)
'Do While Not rs.EOF
'  vNumAsiento = "AF-C{" & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") & "}J" & Trim(rs!cod_justificacion)
'  vAnotacion = "ASIENTO DEPRECIACION RESUMEN X JUSTIFICANTES"
'  vDetalle = "DEP." & rs!Descripcion
'  If rs!depreciacion > 0 Then
'    Call sbGAsientoMaestro("AS", vNumAsiento, vFecha, vDescripcion, vAnotacion, "C")
'    Call sbGAsientoDetalle(1, "AS", vNumAsiento, rs!CG, rs!depreciacion, "D", "PER." & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") _
'                    , vDetalle, "CIERRE")
'    Call sbGAsientoDetalle(2, "AS", vNumAsiento, rs!CD, rs!depreciacion, "H", "PER." & Year(dtpPeriodo.Value) & "-" & Format(Month(dtpPeriodo.Value), "00") _
'                    , vDetalle, "CIERRE")
'  End If
'  rs.MoveNext
'Loop
'rs.Close
'
'End Sub


Private Sub cmdCierre_Click()
Dim strSQL As String

'1. Revisar que el periodo no este cerrado de lo contrario no procesar
' porque ya fue procesado definitivamente...

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_CierreAuxiliar " & Year(dtpPeriodo.Value) & "," & Month(dtpPeriodo.Value) & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call dtpPeriodo_Change

Me.MousePointer = vbDefault

MsgBox "Mes Cerrado Satisfactoriamente...", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpPeriodo_Change()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select estado from Activos_periodos where anio = " & Year(dtpPeriodo.Value) _
       & " and mes = " & Month(dtpPeriodo.Value)
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  If rs!Estado = "P" Then
      lblX.Caption = "  >>>  Periodo Pendiente   <<<"
      lblX.Tag = "P"
  Else
      lblX.Caption = "   >>>   Periodo Cerrado   <<<"
      lblX.Tag = "C"
  End If
Else
  lblX.Caption = "  >>>  Periodo Pendiente   <<<"
  lblX.Tag = "P"
End If
rs.Close
End Sub

Private Function fxCompras() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaI As Date, vFechaC As Date, vRegistro As Currency

On Error GoTo vError

strSQL = "select RegistroCompras from Activos_parametros"
Call OpenRecordSet(rs, strSQL, 0)
If rs!registroCompras = 1 Then
   fxCompras = True
   Exit Function
End If
rs.Close

'Revisar Parametro
lbl.Caption = "Revisando Registro de Compras de Activos, se encuentren en el auxiliar..."


'Revisa que no existan compras de activos sin registrar en el Auxiliar
vFechaI = CDate(Year(dtpPeriodo.Value) & "/" & Format(Month(dtpPeriodo.Value), "00") & "/01")
vFechaC = DateAdd("d", -1, DateAdd("m", 1, vFechaI))

strSQL = "SELECT D.COD_FACTURA,D.LINEA,D.COD_PROVEEDOR,D.COD_PRODUCTO" _
       & ",D.CANTIDAD,P.DESCRIPCION AS PROVEEDOR,P.DESCRIPCION AS PRODUCTO,E.FECHA" _
       & " FROM CPR_COMPRAS E inner join CPR_COMPRAS_detalle D" _
       & " on E.cod_factura = D.cod_factura and E.cod_proveedor = D.cod_proveedor" _
       & " inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " and P.tipo_producto = 'A'" _
       & " inner join cxp_proveedores X on D.cod_proveedor = X.cod_proveedor" _
       & " WHERE E.FECHA BETWEEN '" & Format(vFechaI, "yyyy/mm/dd") & "' AND '" _
       & Format(vFechaC, "yyyy/mm/dd") & "'"
 Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 vRegistro = fxActivos_RegistroCompras(vFechaI, vFechaC, rs!COD_PROVEEDOR, rs!cod_factura)
 If rs!cantidad > vRegistro Then
   lbl.Caption = "Existen Compras de Activos Sin Registrar..."
   fxCompras = False
   Exit Function
 End If
 rs.MoveNext
Loop
rs.Close

fxCompras = True
lbl.Caption = ""

Exit Function

vError:
  fxCompras = True
  lbl.Caption = ""
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


Private Sub Form_Load()
 vModulo = 36
   
 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 dtpPeriodo.Value = gActivos.Periodo
 
 Call dtpPeriodo_Change
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

