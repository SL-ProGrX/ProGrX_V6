VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFNDCC_GeneraAsientos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación de Asientos Pendientes"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5565
   Icon            =   "CC_GeneraAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenera 
      Caption         =   "&Genera Asiento"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimientos"
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.OptionButton optAsientos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Planilla"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optAsientos 
         Caption         =   "Anulación de Aportes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optAsientos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rendimientos"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton optAsientos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Liquidaciones"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin MSComCtl2.DTPicker dtpFechaInicio 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   165347331
         CurrentDate     =   36356
      End
      Begin VB.CheckBox chkGenera 
         Alignment       =   1  'Right Justify
         Caption         =   "Generar todos los pendientes"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpFechaCorte 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   165347331
         CurrentDate     =   36356
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   2760
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   2760
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "GENERAR POR FECHAS"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Corte"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   2880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblEstatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Width           =   4095
   End
End
Attribute VB_Name = "frmFNDCC_GeneraAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxUltimaLineaAsiento(strTipo As String, strCaso As String, vFecha As Date) As Integer
Dim rs As New ADODB.Recordset, strSQL As String, strNumero_Asiento As String

If strTipo = "PR" Then
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
             & Right(Format(Year(vFecha), "0000"), 2) & "N"
Else
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
End If

strSQL = "Select max(num_linea) as Linea from cntx_asientos_detalle" _
       & " where num_asiento = '" & strNumero_Asiento & "' and Tipo_asiento = 'AS'" _
       & " and cod_contabilidad = " & GLOBALES.gEnlace

rs.Open strSQL, glogon.Conection, adOpenStatic

fxUltimaLineaAsiento = IIf(IsNull(rs!Linea), 0, rs!Linea)
rs.Close

End Function


Private Function fxVerificaExistenciaAsiento(strTipo As String, strCaso As String, vFecha As Date) As Boolean
Dim rs As New ADODB.Recordset, strSQL As String, strNumero_Asiento As String

If strTipo = "PR" Then
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
             & Right(Format(Year(vFecha), "0000"), 2) & "N"
Else
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
End If

strSQL = "Select num_asiento from cntx_asientos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
        & " and tipo_asiento = 'AS' and num_asiento = '" & strNumero_Asiento _
        & "' and cod_contabilidad = " & GLOBALES.gEnlace
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF And rs.BOF Then
 fxVerificaExistenciaAsiento = False
Else
 fxVerificaExistenciaAsiento = True
End If
rs.Close
End Function


Private Sub GeneraAsientos(strTipo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim strCasoActual As String, intLinea As Integer, strInforma As String
Dim vFecha As Date, vDetalle  As String


strCasoActual = ""
lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

vFecha = fxFechaServidor

On Error GoTo CapturaError

If chkGenera.Value = vbChecked Then
 strSQL = "select * from fnd_asientos_cola where fnd_fechatrp is null and FND_tipo = '" _
    & strTipo & "' order by fnd_caso,fnd_fecha"

 rs.Open "select coalesce(count(*),0) as Cuenta from fnd_asientos_cola where fnd_fechatrp is null and fnd_tipo = '" _
    & strTipo & "'", glogon.Conection, adOpenStatic
 prgBar.Max = rs!Cuenta + 1
 rs.Close
Else
 strSQL = "select * from fnd_asientos_cola where fnd_fechatrp is null and fnd_tipo = '" _
    & strTipo & "' and fnd_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
    & " 00:00:00' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & " 23:59:59' order by fnd_caso,fnd_fecha"
 
 rs.Open "select coalesce(count(*),0) as Cuenta from fnd_asientos_cola where fnd_fechatrp is null and fnd_tipo = '" _
    & strTipo & "' and fnd_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
    & " 00:00:00' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & " 23:59:59'", glogon.Conection, adOpenStatic
 prgBar.Max = rs!Cuenta + 1
 rs.Close
End If

rs.Open strSQL, glogon.Conection, adOpenStatic

lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh


Do While Not rs.EOF
 If fxgCntPeriodoValida(rs!fnd_Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    
    vDetalle = "TP:" & Trim(rs!fnd_tipo) & " CS:" & Trim(rs!fnd_caso) & " OP:" & rs!cod_operadora & " PLN:" & rs!cod_plan & " CNT:" & rs!cod_contrato
    
    If Trim(rs!fnd_caso) <> strCasoActual Then
        intLinea = 0
        strCasoActual = Trim(rs!fnd_caso)
        
        If fxVerificaExistenciaAsiento(rs!fnd_tipo, rs!fnd_caso, rs!fnd_Fecha) Then
           intLinea = fxUltimaLineaAsiento(rs!fnd_tipo, rs!fnd_caso, rs!fnd_Fecha)
        Else
            Call CreaMaestroAsiento(rs!fnd_tipo, rs!fnd_caso, rs!fnd_Fecha)
            intLinea = 0
        End If
    End If
    
    intLinea = intLinea + 1
    Call CreaDetalleAsiento(rs!fnd_tipo, rs!fnd_caso, IIf(IsNull(rs!fnd_cuenta), "", rs!fnd_cuenta) _
            , rs!fnd_Fecha, rs!fnd_monto, rs!fnd_DEBEHABER, intLinea, vDetalle)
    'Actualizar el estado del asiento
    strSQL = "Update fnd_asientos_cola set fnd_FechaTRP = '" & Format(vFecha, "yyyy/mm/dd") & "' where fnd_asientocola = " & rs!fnd_asientocola
    glogon.Conection.Execute strSQL
 Else
  strInforma = "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado..."
 End If 'Periodo
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

lblEstatus.Caption = ""
lblEstatus.Refresh
prgBar.Value = 1

MsgBox "Se realizó el pase de asientos a contabilidad - " & strInforma, vbInformation
'Bitacora
Call Bitacora("Aplica", "Asientos FND del Tipo Siguiente : " & strTipo)
Exit Sub

CapturaError:
    lblEstatus.Caption = ""
    MsgBox Err.Description, vbCritical
    lblEstatus.Refresh
    prgBar.Value = 1
    rs.Close

End Sub


Public Function fxExisteAsientoNuevo(strTipo As String, vFecha As Date) As Boolean
Dim rs As New ADODB.Recordset, strSQL As String, strNumero_Asiento As String

strNumero_Asiento = "FND-" & strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "Select * from asientos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
        & " and tipo_asiento = 'AS' and num_asiento = '" & strNumero_Asiento & "'"

rs.Open strSQL, glogon.Conection, adOpenStatic

If rs.EOF And rs.BOF Then
 fxExisteAsientoNuevo = False
Else
 fxExisteAsientoNuevo = True
End If
rs.Close
End Function

Public Function fxUltimaLineaAsientoNuevo(strTipo As String, vFecha As Date) As Integer
Dim rs As New ADODB.Recordset, strSQL As String, strNumero_Asiento As String

strNumero_Asiento = "FND-" & strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "Select max(num_linea) as Linea from asientos_detalle where num_asiento = '" & strNumero_Asiento & "'"

rs.Open strSQL, glogon.Conection, adOpenStatic

fxUltimaLineaAsientoNuevo = IIf(IsNull(rs!Linea), 0, rs!Linea)
rs.Close

End Function

''Sub GeneraAsientosFechas(strTipo As String)
''Dim rs As New ADODB.Recordset, strSQL As String
''Dim intLinea As Integer, strInforma As String
''Dim vFecha As Date, vUltFecha As Date, vNumeroAsiento As String
''
''lblEstatus.Caption = "Cargando Información..."
''lblEstatus.Refresh
''prgBar.Value = 1
''
''vFecha = fxFechaServidor
''vUltFecha = "1980/01/01" 'Fecha de Inicio
''
''
''On Error GoTo vError
''
''If chkGenera.Value = vbChecked Then
'' strSQL = "select * from fnd_asientos_cola where fnd_fechatrp is null and fnd_tipo = '" _
''    & strTipo & "' order by fnd_fecha,fnd_caso"
''
'' rs.Open strSQL, glogon.Conection, adOpenStatic
'' prgBar.Max = rs.RecordCount + 1
'' rs.Close
''
''Else
''
'' strSQL = "select * from fnd_asientos_cola where fnd_fechatrp is null and fnd_tipo = '" _
''    & strTipo & "' and fnd_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
''    & "' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "' order by fnd_fecha,fnd_caso"
'' rs.Open strSQL, glogon.Conection, adOpenStatic
'' prgBar.Max = rs.RecordCount + 1
'' rs.Close
''End If
''
''rs.Open strSQL, glogon.Conection, adOpenStatic
''
''lblEstatus.Caption = "Procesando Asientos..."
''lblEstatus.Refresh
''
''
''Do While Not rs.EOF
'' If fxValidaPeriodoAsiento(rs!fnd_Fecha) Then 'Verificar el Periodo Abierto en contabilidad
''    If rs!fnd_Fecha <> vUltFecha Then
''
''     intLinea = 0
''     vUltFecha = rs!fnd_Fecha
''     vNumeroAsiento = "FND-" & strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
''
''     If fxExisteAsientoNuevo(rs!fnd_tipo, rs!fnd_Fecha) Then
''        intLinea = fxUltimaLineaAsientoNuevo(rs!fnd_tipo, rs!fnd_Fecha)
''     Else
''         Call CreaMaestroAsientoNuevo(rs!fnd_tipo, rs!fnd_Fecha)
''     End If
''    End If
''    intLinea = intLinea + 1
''    Call CreaDetalleAsientoNuevo(rs!fnd_tipo, IIf(IsNull(rs!fnd_cuenta), "", rs!fnd_cuenta), rs!fnd_Fecha, rs!fnd_Monto, rs!fnd_DEBEHABER, intLinea, rs!fnd_caso)
''    'Actualizar el estado del asiento
''    strSQL = "Update fnd_asientos_cola set fnd_FechaTRP = '" & Format(vFecha, "yyyy/mm/dd") & "' where fnd_asientocola = " & rs!fnd_asiento_cola
''    glogon.Conection.Execute strSQL
'' Else
''  strInforma = "Existen asientos que no pueden ser trasladados porque el periodo fué cerrado..."
'' End If 'Periodo
'' prgBar.Value = prgBar.Value + 1
'' rs.MoveNext
''Loop
''rs.Close
''
''lblEstatus.Caption = ""
''lblEstatus.Refresh
''prgBar.Value = 1
''
'''Bitacora
''Call Bitacora("Aplica", "Asientos del Tipo Siguiente : " & strTipo)
''
''MsgBox "Se realizó el pase de asientos a contabilidad - " & strInforma, vbInformation
''
''Exit Sub
''
''vError:
''    lblEstatus.Caption = ""
''    MsgBox Err.Description, vbCritical
''    lblEstatus.Refresh
''    prgBar.Value = 1
''    rs.Close
''
''End Sub

Private Sub chkGenera_Click()
If chkGenera.Value = vbChecked Then
 dtpFechaInicio.Enabled = False
 dtpFechaCorte.Enabled = False
Else
 dtpFechaInicio.Enabled = True
 dtpFechaCorte.Enabled = True
End If
End Sub

Private Sub cmdGenera_Click()
Dim iRespuesta As Integer
 iRespuesta = MsgBox("Esta seguro que desea Generar Asientos a Contabilidad", vbYesNo)
 If iRespuesta = vbYes Then
  Select Case True
    Case optAsientos(0).Value 'Liquidacion
      Call GeneraAsientos("LI")
    Case optAsientos(1).Value 'Rendimientos
      Call GeneraAsientos("CR")
    Case optAsientos(2).Value 'Planillas
      Call GeneraAsientos("PR")
    Case Else
     'No esta definido
  End Select
 End If
End Sub

Sub CreaMaestroAsientoNuevo(strTipo As String, vFecha As Date)
Dim strSQL As String, strNumero_Asiento As String

On Error GoTo ErrorCap
     
strNumero_Asiento = "FND-" & strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "insert asientos(tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,anotacion,balanceado)" _
    & " values('AS','" & strNumero_Asiento & "'," & Year(vFecha) & "," & Month(vFecha) _
    & ",'" & Format(vFecha, "yyyy/mm/dd") & "','TRASPASO DE ASIENTOS DEL FND','DIARIO','S')"
glogon.Conection.Execute strSQL

Exit Sub

ErrorCap:
MsgBox Err.Description, vbCritical

End Sub

''Sub CreaDetalleAsientoNuevo(strTipo As String, strCuenta As String, vFecha As Date _
''                          , curMonto As Currency, DH As String, intLinea As Integer, Operacion As Long)
''Dim strSQL As String, strNumero_Asiento As String
''Dim rs As New ADODB.Recordset, vDetalle
''
''If UCase(DH) <> "D" Then 'dc - dh
''  DH = "C"
''End If
''
''strSQL = "select nombre from socios S inner join reg_creditos R on S.cedula = R.cedula " _
''       & "where R.id_solicitud = " & Operacion
''rs.Open strSQL, glogon.Conection, adOpenStatic
''If rs.BOF And rs.EOF Then
'' 'no deberia de ocurrir
''  vDetalle = ""
''Else
''  vDetalle = Operacion & "-" & rs!Nombre & ""
''  vDetalle = Mid(vDetalle, 1, 59)
''End If
''rs.Close
''
''strNumero_Asiento = "FND-" & strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
''
''strSQL = "insert asientos_detalle(tipo_asiento,num_asiento,num_linea,num_cuenta,tipo_movimiento,monto,detalle,referencia,num_documento)" _
''    & " values('AS','" & strNumero_Asiento & "'," & intLinea & "," & fxNumeroCuenta(strCuenta) & ",'" & DH & "'," _
''    & curMonto & ",'" & vDetalle & "','TRASPASO-FND','" & Operacion & "')"
''
''glogon.Conection.Execute strSQL
''
''End Sub



Sub CreaMaestroAsiento(strTipo As String, strCaso As String, vFecha As Date)
Dim strSQL As String, strNumero_Asiento As String

On Error GoTo ErrorCap

If strTipo = "PR" Then
  strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
           & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
           & Right(Format(Year(vFecha), "0000"), 2) & "N"
Else
  strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
           & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
End If

strSQL = "insert CntX_asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
       & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & Year(vFecha) & "," & Month(vFecha) _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','TRASPASO DE ASIENTOS DEL FND','S'," & vModulo & " )"
glogon.Conection.Execute strSQL

Exit Sub

ErrorCap:
MsgBox Err.Description, vbCritical


End Sub

Sub CreaDetalleAsiento(strTipo As String, strCaso As String, strCuenta As String, vFecha As Date _
           , curMonto As Currency, DH As String, intLinea As Integer, Optional vDetalle As String = "")
Dim strSQL As String, strNumero_Asiento As String

If UCase(DH) <> "D" Then 'dc - dh
  DH = "C"
End If

If strTipo = "PR" Then
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00") _
             & Right(Format(Year(vFecha), "0000"), 2) & "N"
Else
    strNumero_Asiento = "F" & strTipo & "-" & Format(Mid(Trim(strCaso), 1, 5), "00000#") _
             & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
End If

If vDetalle = "" Then vDetalle = strCaso

If DH = "C" Then  'Acredita
   strSQL = "insert CntX_asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,tipo_cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & "," & Trim(strCuenta) & ",0," _
          & curMonto & ",'" & vDetalle & "','" & strTipo & "','OC','COL',1,'')"
Else 'Debita
   strSQL = "insert CntX_asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,tipo_cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & "," & Trim(strCuenta) & "," _
          & curMonto & ",0,'" & vDetalle & "','" & strTipo & "','OC','COL',1,'')"
End If

glogon.Conection.Execute strSQL

End Sub

Private Sub Form_Load()
 vModulo = 18
 
 dtpFechaInicio.Value = Date
 dtpFechaCorte.Value = Date
 
' Call Formularios(Me)
' Call RefrescaTags(Me)
End Sub


