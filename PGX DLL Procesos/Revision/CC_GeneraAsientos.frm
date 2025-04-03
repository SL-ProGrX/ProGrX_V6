VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCC_GeneraAsientos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generación de Asientos Pendientes"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5175
   Icon            =   "CC_GeneraAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formalizaciones Anuladas"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   8
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox chkGenera 
      Caption         =   "Todas"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdGenera 
      Caption         =   "&Generar"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   4650
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpFechaInicio 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   110297091
      CurrentDate     =   36356
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   110297091
      CurrentDate     =   36356
   End
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formalizaciones"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   7
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Planillas (Patrimonio/Crédito)"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   6
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   2415
   End
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Traslado de Cuentas (Cobro)"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresos (Afiliación)"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   2415
   End
   Begin VB.OptionButton optAsientos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liquidaciones Socios"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipos de Asientos en Cola"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Parámetros"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblEstatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   3735
   End
End
Attribute VB_Name = "frmCC_GeneraAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub sbGeneraAsientos(strTipo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim strCasoActual As String, intLinea As Integer, strInforma As String
Dim vFecha As Date, vNumAsiento As String

strCasoActual = ""
lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

vFecha = fxFechaServidor

On Error GoTo vError

If chkGenera.Value = vbChecked Then
 strSQL = "select * from asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "' order by tmp_caso"

 rs.Open "select count(*) as Cuenta from asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "'", glogon.Conection, adOpenStatic
 prgBar.Max = IIf(IsNull(rs!Cuenta), 1, rs!Cuenta) + 1
 rs.Close
Else
 strSQL = "select * from asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "' and tmp_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
    & "' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "' order by tmp_fecha,tmp_caso"
 
 rs.Open "select coalesce(count(*),0) + 1 as Cuenta from asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "' and tmp_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
    & "' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "'", glogon.Conection, adOpenStatic
 prgBar.Max = rs!Cuenta
 rs.Close
End If

rs.Open strSQL, glogon.Conection, adOpenStatic

lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh


Do While Not rs.EOF
 If fxValidaPeriodoAsiento(rs!tmp_Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    
    If rs!tmp_tipo = "TRA" Then
        vNumAsiento = Trim(rs!tmp_tipo) & "-" & Trim(rs!tmp_caso) & "-Y" & Right(Format(Year(vFecha), "0000"), 2)
    Else
        vNumAsiento = Trim(rs!tmp_tipo) & "-" & Format(Mid(Trim(rs!tmp_caso), 1, 5), "00000#") _
                 & "D-" & Format(rs!tmp_Fecha, "yyyymmdd")
    End If
    
    If Trim(rs!tmp_caso) <> strCasoActual Then
     intLinea = 0
     strCasoActual = Trim(rs!tmp_caso)
     
     If fxVerificaExistenciaAsiento(vNumAsiento, rs!tmp_Fecha) Then
        intLinea = fxUltimaLineaAsiento(vNumAsiento, rs!tmp_Fecha)
     Else
         Call sbAsientoMaestro(vNumAsiento, rs!tmp_Fecha)
     End If
    
   End If ' tmp_Caso
    
    intLinea = intLinea + 1
    Call sbAsientoDetalle(vNumAsiento, Trim(rs!tmp_tipo), Trim(rs!tmp_caso), IIf(IsNull(rs!tmp_cuenta), "" _
            , rs!tmp_cuenta), rs!tmp_Fecha, rs!tmp_Monto, rs!TMP_DEBEHABER, intLinea)
    
    'Actualizar el estado del asiento
    strSQL = "Update asientos_tmp set tmp_FechaTRP = '" & Format(vFecha, "yyyy/mm/dd") & "' where tmp_id_asiento = " & rs!tmp_id_asiento
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
Call Bitacora("Aplica", "Asientos del Tipo Siguiente : " & strTipo)
Exit Sub

vError:
    lblEstatus.Caption = ""
    MsgBox Err.Description, vbCritical
    lblEstatus.Refresh
    prgBar.Value = 1
    rs.Close

End Sub


Public Function fxExisteAsientoNuevo(strTipo As String, vFecha As Date) As Boolean
Dim rs As New ADODB.Recordset, strSQL As String, strNumero_Asiento As String

strNumero_Asiento = strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "Select * from CntX_Asientos where anio = " & Year(vFecha) & " and mes = " & Month(vFecha) _
        & " and tipo_asiento = 'AS' and num_asiento = '" & strNumero_Asiento _
        & "' and cod_contabilidad = " & GLOBALES.gEnlace
 
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

strNumero_Asiento = strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "Select max(num_linea) as Linea from CntX_Asientos_detalle where num_asiento = '" _
       & strNumero_Asiento & "' and tipo_asiento = 'AS' and cod_contabilidad = " & GLOBALES.gEnlace

rs.Open strSQL, glogon.Conection, adOpenStatic

fxUltimaLineaAsientoNuevo = IIf(IsNull(rs!linea), 0, rs!linea)
rs.Close

End Function

Private Sub sbsbGeneraAsientosFechas(strTipo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim intLinea As Integer, strInforma As String
Dim vFecha As Date, vUltFecha As Date, vNumeroAsiento As String

lblEstatus.Caption = "Cargando Información..."
lblEstatus.Refresh
prgBar.Value = 1

vFecha = fxFechaServidor
vUltFecha = "1980/01/01" 'Fecha de Inicio


On Error GoTo vError

If chkGenera.Value = vbChecked Then
 strSQL = "select * from Asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "' order by tmp_fecha,tmp_caso"
 
 rs.Open strSQL, glogon.Conection, adOpenStatic
 prgBar.Max = rs.RecordCount + 1
 rs.Close

Else

 strSQL = "select * from Asientos_tmp where tmp_fechatrp is null and tmp_tipo = '" _
    & strTipo & "' and tmp_fecha between '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
    & "' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "' order by tmp_fecha,tmp_caso"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 prgBar.Max = rs.RecordCount + 1
 rs.Close
End If

rs.Open strSQL, glogon.Conection, adOpenStatic

lblEstatus.Caption = "Procesando Asientos..."
lblEstatus.Refresh


Do While Not rs.EOF
 If fxValidaPeriodoAsiento(rs!tmp_Fecha) Then 'Verificar el Periodo Abierto en contabilidad
    If rs!tmp_Fecha <> vUltFecha Then
     
     intLinea = 0
     vUltFecha = rs!tmp_Fecha
     vNumeroAsiento = strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
     
     If fxExisteAsientoNuevo(rs!tmp_tipo, rs!tmp_Fecha) Then
        intLinea = fxUltimaLineaAsientoNuevo(rs!tmp_tipo, rs!tmp_Fecha)
     Else
         Call sbAsientoMaestroFechas(rs!tmp_tipo, rs!tmp_Fecha)
     End If
    End If
    intLinea = intLinea + 1
    Call sbAsientoDetalleFechas(rs!tmp_tipo, IIf(IsNull(rs!tmp_cuenta), "", rs!tmp_cuenta), rs!tmp_Fecha, rs!tmp_Monto, rs!TMP_DEBEHABER, intLinea, rs!tmp_operacion)
    'Actualizar el estado del asiento
    strSQL = "Update asientos_tmp set tmp_FechaTRP = '" & Format(vFecha, "yyyy/mm/dd") & "' where tmp_id_asiento = " & rs!tmp_id_asiento
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

'Bitacora
Call Bitacora("Aplica", "Asientos del Tipo Siguiente : " & strTipo)

MsgBox "Se realizó el pase de asientos a contabilidad - " & strInforma, vbInformation

Exit Sub

vError:
    lblEstatus.Caption = ""
    MsgBox Err.Description, vbCritical
    lblEstatus.Refresh
    prgBar.Value = 1
    rs.Close

End Sub

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
Dim i As Byte

i = MsgBox("Esta seguro que desea Generar Asientos a Contabilidad", vbYesNo)
If i = vbNo Then Exit Sub
 
Select Case True
  Case optAsientos(0).Value 'Liquidacion Socios
   Call sbGeneraAsientos("LIQ")
  Case optAsientos(1).Value 'Reingresos e Ingreso
   Call sbGeneraAsientos("ING")
  Case optAsientos(4).Value 'Traslados de cuentas COBRO
   Call sbGeneraAsientos("TRA")
  Case optAsientos(6).Value 'Proceso Mensual
   Call sbGeneraAsientos("PRM")
  Case optAsientos(7).Value 'Formalizaciones
   Call sbsbGeneraAsientosFechas("FRM")
  Case optAsientos(8).Value 'Formalizaciones Anuladas
   Call sbsbGeneraAsientosFechas("AFR")
End Select

End Sub

Sub sbAsientoMaestroFechas(strTipo As String, vFecha As Date)
Dim strSQL As String, strNumero_Asiento As String

On Error GoTo ErrorCap
     
strNumero_Asiento = strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")

strSQL = "insert cntx_asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado)" _
    & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & Year(vFecha) & "," & Month(vFecha) _
    & ",'" & Format(vFecha, "yyyy/mm/dd") & "','TRASPASO DE ASIENTOS DEL SIF','S')"
glogon.Conection.Execute strSQL

Exit Sub

ErrorCap:
MsgBox Err.Description, vbCritical

End Sub

Private Sub sbAsientoDetalleFechas(strTipo As String, strCuenta As String, vFecha As Date _
                          , curMonto As Currency, DH As String, intLinea As Integer, Operacion As Long)
Dim strSQL As String, strNumero_Asiento As String
Dim rs As New ADODB.Recordset, vDetalle

If UCase(DH) <> "D" Then 'dc - dh
  DH = "C"
End If

strSQL = "select nombre from socios S inner join reg_creditos R on S.cedula = R.cedula " _
       & "where R.id_solicitud = " & Operacion
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.BOF And rs.EOF Then
 'no deberia de ocurrir
  vDetalle = ""
Else
  vDetalle = Operacion & "-" & rs!Nombre & ""
  vDetalle = Mid(vDetalle, 1, 59)
End If
rs.Close

strNumero_Asiento = strTipo & "-" & Format(Year(vFecha), "0000") & Format(Month(vFecha), "00") & Format(Day(vFecha), "00")
If DH = "C" Then 'Acredita
   strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & ",'" & Trim(strCuenta) & "',0," _
          & curMonto & ",'" & vDetalle & "','" & Operacion & "','OC','COL',1,'')"
Else 'Debita
   strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & strNumero_Asiento & "'," & intLinea & ",'" & Trim(strCuenta) & "'," _
          & curMonto & ",0,'" & vDetalle & "','" & Operacion & "','OC','COL',1,'')"
End If

glogon.Conection.Execute strSQL

End Sub



Private Sub sbAsientoMaestro(xNumAsiento As String, vFecha As Date)
Dim strSQL As String

On Error GoTo vError

strSQL = "insert CntX_Asientos(cod_contabilidad,tipo_asiento,num_asiento,anio,mes,fecha_asiento,descripcion,balanceado,modulo)" _
    & " values(" & GLOBALES.gEnlace & ",'AS','" & xNumAsiento & "'," & Year(vFecha) & "," & Month(vFecha) _
    & ",'" & Format(vFecha, "yyyy/mm/dd") & "','TRASPASO DE ASIENTOS DEL SIF','S'," & vModulo & ")"
glogon.Conection.Execute strSQL

Exit Sub

vError:
MsgBox Err.Description, vbCritical
End Sub

Private Sub sbAsientoDetalle(xNumAsiento As String, strTipo As String, strCaso As String, strCuenta As String _
                , vFecha As Date, curMonto As Currency, DH As String, intLinea As Integer)
Dim strSQL As String

If UCase(DH) <> "D" Then DH = "C"

If DH = "C" Then  'Acredita
   strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & xNumAsiento & "'," & intLinea & "," & Trim(strCuenta) & ",0," _
          & curMonto & ",'" & strCaso & "','" & strTipo & "','OC','COL',1,'')"
Else 'Debita
   strSQL = "insert CntX_Asientos_detalle(cod_contabilidad,tipo_asiento,num_asiento,num_linea,cod_cuenta,monto_debito,monto_credito" _
          & ",detalle,documento,cod_unidad,cod_divisa,Tipo_Cambio,cod_centro_costo)" _
          & " values(" & GLOBALES.gEnlace & ",'AS','" & xNumAsiento & "'," & intLinea & "," & Trim(strCuenta) & "," _
          & curMonto & ",0,'" & strCaso & "','" & strTipo & "','OC','COL',1,'')"
End If

glogon.Conection.Execute strSQL

End Sub

Private Sub Form_Load()
 dtpFechaInicio.Value = Date
 dtpFechaCorte.Value = Date
 
 vModulo = 10 'Cuentas Corrientes
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub
