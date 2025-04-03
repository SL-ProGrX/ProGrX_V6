VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMigracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migra Contabilidad del Sistema ASE a ContaExpress"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblEstado 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Parametros
  Nivel1  As Integer
  Nivel2  As Integer
  Nivel3  As Integer
  Nivel4  As Integer
  Nivel5  As Integer
End Type
Dim vParametros As Parametros

Private Function fxFechaAsiento(vNum As String, vTipo As String, AdoC As ADODB.Connection) As Date
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select fecha_asiento from asientos where num_asiento = '" & vNum _
        & "' and tipo_asiento = '" & vTipo & "'"
rsX.CursorLocation = adUseServer
rsX.Open strSQL, AdoC, adOpenStatic

If rsX.EOF And rsX.BOF Then
 fxFechaAsiento = Date
 Exit Function
Else
fxFechaAsiento = rsX!fecha_asiento
End If

rsX.Close

End Function

Private Function fxExisteCuenta(lngEmpresa As Long, lngAnio As Long, iMes As Integer, vCuenta As String, AdoC As ADODB.Connection) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select * from movimiento_cuentas where anio = " & lngAnio _
       & " and mes = " & iMes & " and cod_empresa = " & lngEmpresa _
       & " and cod_cuenta = '" & vCuenta & "'"
rsX.Open strSQL, AdoC, adOpenStatic

If rsX.EOF And rsX.BOF Then
 fxExisteCuenta = False
Else
 fxExisteCuenta = True
End If

rsX.Close

End Function


Private Function fxCuentaMadre(AdoC As ADODB.Connection, vCuenta As Long) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select cod_cuenta from cuentas where num_cuenta = " & vCuenta
rsX.CursorLocation = adUseServer

If vCuenta = 0 Then
 fxCuentaMadre = ""
 Exit Function
End If
rsX.Open strSQL, AdoC, adOpenStatic
fxCuentaMadre = rsX!cod_cuenta
rsX.Close

End Function


Private Function fxNivelCuenta(vCuentaMadre As String) As Integer

fxNivelCuenta = 0

If vCuentaMadre = "" Then
  fxNivelCuenta = 1
  Exit Function
End If

With vParametros
  If .Nivel1 > 0 And Val(Mid(vCuentaMadre, 1, .Nivel1)) > 0 Then fxNivelCuenta = 2
  If .Nivel2 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + 1, .Nivel1)) > 0 Then fxNivelCuenta = 3
  If .Nivel3 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + .Nivel2 + 1, .Nivel3)) > 0 Then fxNivelCuenta = 4
  If .Nivel4 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + .Nivel2 + .Nivel3 + 1, .Nivel4)) > 0 Then fxNivelCuenta = 5
End With

If fxNivelCuenta = 0 Then fxNivelCuenta = 5

' MsgBox fxNivelCuenta

End Function


Private Sub cmdAplicar_Click()
Dim AdoConection As New ADODB.Connection, rs As New ADODB.Recordset
Dim AdoConection2 As New ADODB.Connection, strSQL As String
Dim lngEmpresa As Long, vCuenta As String

On Error GoTo vError

strSQL = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\ContaExpress.mdb;Mode=ReadWrite;pwd=xrjk2if9k"""
'Cambiada x Esta
strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=in_pedro;UID=sa;PWD=;Database=ContaExpress;"
AdoConection.Open strSQL

strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=perseus3;UID=sa;PWD=;Database=aseccss;"
AdoConection2.Open strSQL

With vParametros
  .Nivel1 = 3
  .Nivel2 = 2
  .Nivel3 = 3
  .Nivel4 = 2
  .Nivel5 = 0
End With

strSQL = "insert into empresas(nombre,nivel1,nivel2,nivel3,nivel4,nivel5) values(" _
        & "'ASECCSS ASE',3,2,3,2,0)"
AdoConection.Execute strSQL

lngEmpresa = 1

lblEstado.Caption = "Procesando Periodos"
lblEstado.Refresh
 
strSQL = "select * from periodos"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "insert into periodos(cod_empresa,anio,mes,estado) values(" & lngEmpresa & "," & rs!anio _
  & "," & rs!mes & ",'" & IIf((rs!estado = "S"), "C", "P") & "')"
  AdoConection.Execute strSQL
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
      
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Procesando Tipos de Cuentas"
lblEstado.Refresh


strSQL = "select * from tipos_cuenta"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "insert into tipos_cuentas(cod_empresa,tipo_cuenta,clasificacion,descripcion) values(" & lngEmpresa _
        & ",'" & rs!tipo_cuenta & "','" & rs!clasificacion & "','" & rs!descripcion & "')"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Procesando Tipos de Asientos"
lblEstado.Refresh

strSQL = "select * from tipos_asiento"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1


Do While Not rs.EOF
  strSQL = "insert into tipos_asientos(cod_empresa,tipo_asiento,descripcion) values(" & lngEmpresa _
        & ",'" & rs!tipo_asiento & "','" & rs!descripcion & "')"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Procesando Catalogo de Cuentas"
lblEstado.Refresh

strSQL = "select * from cuentas"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "insert into cuentas(cod_empresa,tipo_cuenta,descripcion,cod_cuenta,cuenta_madre," _
        & "acepta_movimientos) values(" & lngEmpresa & ",'" & rs!tipo_cuenta & "','" & rs!descripcion _
        & "','" & rs!cod_cuenta & "','" & fxCuentaMadre(AdoConection2, IIf(IsNull(rs!num_cuenta_madre), 0, rs!num_cuenta_madre)) _
        & "','" & rs!acepta_movimientos & "')"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Procesando Movimientos de Cuentas"
lblEstado.Refresh

strSQL = "select * from movimientos_cuentas"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "insert into movimiento_cuentas(cod_empresa,cod_cuenta,anio,mes,saldo_inicial," _
        & "total_debitos,total_creditos) values(" & lngEmpresa & ",'" & fxCuentaMadre(AdoConection2, rs!num_cuenta) _
        & "'," & rs!anio & "," & rs!mes & "," & rs!saldo_inicial & "," & rs!total_debitos _
        & "," & rs!total_creditos & ")"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close



'
'lblEstado.Caption = "Reestructurando Movimientos de Cuentas"
'lblEstado.Refresh
'
'AdoConection.CommandTimeout = 1000
'
'strSQL = "select M.*,C.cuenta_madre from movimiento_cuentas M" _
'       & " inner join Cuentas C on C.cod_cuenta = M.cod_cuenta" _
'       & " where M.cod_empresa = C.cod_empresa and C.cod_empresa = " & lngEmpresa _
'       & " and C.acepta_movimientos = 'N'"
'rs.Open strSQL, AdoConection, adOpenStatic 'En Conta Express
'prgBar.Value = 1
'prgBar.Max = rs.RecordCount + 1
'
'lblEstado.Caption = "Limpiando Cuentas Invalidas"
'lblEstado.Refresh
'
'Do While Not rs.EOF
'  strSQL = "delete * from movimiento_cuentas where cod_cuenta = '" _
'         & rs!cod_cuenta & "' and cod_empresa = " & lngEmpresa
'  AdoConection.Execute strSQL
'  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
'  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
'  lblTotal.Refresh
'  rs.MoveNext
'Loop
'rs.Close
'
'lblEstado.Caption = "Reestructurando Movimientos de Cuentas"
'lblEstado.Refresh
'
'strSQL = "select M.*,C.cuenta_madre from movimiento_cuentas M" _
'       & " inner join Cuentas C on C.cod_cuenta = M.cod_cuenta" _
'       & " where M.cod_empresa = C.cod_empresa and C.cod_empresa = " & lngEmpresa _
'       & " and C.acepta_movimientos = 'S' and C.cuenta_madre <> ''"
'
'rs.Open strSQL, AdoConection, adOpenStatic 'En Conta Express
'
'prgBar.Value = 1
'prgBar.Max = rs.RecordCount + 1
'
'Do While Not rs.EOF
'  If fxExisteCuenta(lngEmpresa, rs!anio, rs!mes, rs!cuenta_madre, AdoConection) Then
'   strSQL = "update movimiento_cuentas set saldo_inicial = saldo_inicial + " & rs!saldo_inicial _
'          & ", total_debitos = total_debitos + " & rs!total_debitos _
'          & ", total_creditos = total_creditos + " & rs!total_creditos _
'          & " where anio = " & rs!anio & " and Mes = " & rs!mes _
'          & " and cod_cuenta = '" & rs!cuenta_madre & "' and cod_empresa = " & lngEmpresa
'  Else
'  strSQL = "insert into movimiento_cuentas(cod_empresa,cod_cuenta,anio,mes,saldo_inicial," _
'        & "total_debitos,total_creditos) values(" & lngEmpresa & ",'" & rs!cuenta_madre _
'        & "'," & rs!anio & "," & rs!mes & "," & rs!saldo_inicial & "," & rs!total_debitos _
'        & "," & rs!total_creditos & ")"
'  End If
'  AdoConection.Execute strSQL
'
'  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
'  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
'  lblTotal.Refresh
'  rs.MoveNext
'Loop
'rs.Close
'


lblEstado.Caption = "Procesando Maestro de Asientos"
lblEstado.Refresh

strSQL = "select * from asientos"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "insert into asientos(cod_empresa,anio,mes,num_asiento,tipo_asiento,fecha_asiento," _
        & "descripcion,balanceado,fecha_aplicado) values(" & lngEmpresa & "," & rs!anio _
        & "," & rs!mes & ",'" & rs!num_asiento & "','" & rs!tipo_asiento _
        & "','" & Format(rs!fecha_asiento, "yyyy/mm/dd") & "','" & Mid(rs!descripcion, 1, 54) & "','" _
        & rs!balanceado & "','" & Format(rs!fecha_aplicado, "yyyy/mm/dd") & "')"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Procesando Detalles de Asientos"
lblEstado.Refresh

AdoConection2.CommandTimeout = 1000

strSQL = "select D.*,A.fecha_asiento,C.cod_cuenta" _
       & " from asientos_detalle D inner join asientos A" _
       & " on D.num_asiento = A.num_asiento and D.tipo_asiento = A.tipo_asiento" _
       & " inner join cuentas C on D.num_cuenta = C.num_cuenta"
rs.Open strSQL, AdoConection2, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

lblTotal.Width = lblTotal.Width * 2

Do While Not rs.EOF
  strSQL = "insert into asientos_detalle(cod_empresa,num_asiento,tipo_asiento,num_linea," _
        & "cod_cuenta,documento,detalle,monto_debito,monto_credito,fecha_asiento) values(" _
        & lngEmpresa & ",'" & rs!num_asiento & "','" & rs!tipo_asiento & "'," & rs!num_linea _
        & ",'" & Trim(rs!cod_cuenta) & "','" & rs!num_documento _
        & "','" & rs!detalle & "'," & IIf((rs!tipo_movimiento = "D"), rs!monto, 0) _
        & "," & IIf((rs!tipo_movimiento = "D"), 0, rs!monto) & ",'" _
        & Format(rs!fecha_asiento, "yyyy/mm/dd") & "')"
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Detallando Niveles de Cuentas"
lblEstado.Refresh

strSQL = "select * from cuentas"
rs.Open strSQL, AdoConection, adOpenStatic

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1

Do While Not rs.EOF
  strSQL = "Update cuentas set nivel = " & fxNivelCuenta(rs!cuenta_madre) _
         & " where cod_cuenta = '" & Trim(rs!cod_cuenta) & "' and cod_empresa = " & lngEmpresa
  AdoConection.Execute strSQL
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  lblTotal.Caption = "Procesando " & prgBar.Value & " de " & prgBar.Max
  lblTotal.Refresh
  rs.MoveNext
Loop
rs.Close


AdoConection.Close
AdoConection2.Close

MsgBox "Migración Terminada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 Resume

End Sub

