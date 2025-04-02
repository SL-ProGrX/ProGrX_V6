VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFNDCierreMensualFND 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierre Mensual del Fondo"
   ClientHeight    =   3048
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9852
   HelpContextID   =   7001
   Icon            =   "frmFNDCierreMensualFND.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3048
   ScaleWidth      =   9852
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   2892
      Width           =   9852
      _ExtentX        =   17378
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdCierre 
      Height          =   852
      Left            =   7440
      TabIndex        =   3
      Top             =   1440
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Cierre Manual de Auxiliar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmFNDCierreMensualFND.frx":08CA
   End
   Begin XtremeSuiteControls.Label lbl 
      Height          =   1212
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   6852
      _Version        =   1245187
      _ExtentX        =   12086
      _ExtentY        =   2138
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cierre de Auxiliar de Planes de Ahorros"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmFNDCierreMensualFND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxExisteCuenta(iMes As Integer, lngAnio As Long, vCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset

rsX.Open "select isnull(count(*),0) as existe from fnd_per_cuentas where anio = " _
        & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
        & Trim(vCuenta) & "'", glogon.Conection, adOpenStatic
If rsX!Existe = 0 Then
  fxExisteCuenta = False
Else
  fxExisteCuenta = True
End If
rsX.Close

End Function


Private Sub sbActualizaMes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String

'***********************************************************************************
'Este codigo no se utiliza, se realizó simplemente para iniciar los nuevos procesos.
'***********************************************************************************

strSQL = "select A.codigo,proceso,opex,sum(saldo_final) as SF,C.ctanamort,C.ctaoamort,C.ctacamort" _
       & " From ase_per_cerrados A inner join catalogo C on A.codigo = C.codigo" _
       & " Where anio = 2001 And mes = 6 group by A.codigo,proceso,opex,C.ctanamort,C.ctaoamort,C.ctacamort"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case rs!Proceso
   Case "J"
       vCuenta = rs!ctacamort
   Case Else
     If rs!opex = 1 Then
       vCuenta = rs!CtaOamort
     Else
       vCuenta = rs!CtaNamort
     End If
 End Select
  
 If rs!sf > 0 Then
    If fxExisteCuenta(7, 2001, vCuenta) Then
          strSQL = "update ase_per_cuentas set Saldo_inicial = Saldo_inicial + " & rs!sf _
                 & " where anio = 2001 and mes = 7 and cod_cuenta = '" & Trim(vCuenta) & "'"
    Else
        strSQL = "insert into ase_per_cuentas(anio,mes,cod_cuenta,saldo_inicial,total_debitos,total_creditos,saldo_final)" _
               & " values(2001,7,'" & Trim(vCuenta) & "'," & rs!sf & ",0,0,0)"
    End If
    Call ConectionExecute(strSQL)
 End If
 
 rs.MoveNext

Loop
rs.Close

MsgBox "Ok"
End Sub

Private Sub sbMovCuentas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iMes As Integer, lngAnio As Long
Dim iMesN As Integer, lngAnioN As Long
Me.MousePointer = vbHourglass


iMes = Month(fxFechaServidor)
lngAnio = Year(fxFechaServidor)

lbl.Caption = "Procesando Resumen de Movimientos de Cuentas (I)..."
lbl.Refresh

strSQL = "select A.fnd_cuenta,A.fnd_debehaber,isnull(sum(A.fnd_monto),0) as Movimiento" _
       & " from fnd_documentos D inner join fnd_asientos A on D.tipo = A.tipo" _
       & " and D.id_documento = A.id_documento and D.cod_operadora = A.cod_operadora" _
       & " where datepart(yyyy, D.fecha) = " & lngAnio _
       & " and datepart(mm, D.fecha) = " & iMes & " group by A.fnd_cuenta,A.fnd_debehaber"

Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   If fxExisteCuenta(iMes, lngAnio, rs!fnd_cuenta) Then
      'Existe
      If rs!fnd_DEBEHABER = "D" Then
        strSQL = "update fnd_per_cuentas set total_debitos = total_debitos + " & rs!Movimiento _
               & " where anio = " & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
               & Trim(rs!fnd_cuenta) & "'"
      Else
        strSQL = "update fnd_per_cuentas set total_creditos = total_creditos + " & rs!Movimiento _
               & " where anio = " & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
               & Trim(rs!fnd_cuenta) & "'"
      End If
   
   Else 'No existe la cuenta
       strSQL = "insert fnd_per_cuentas(anio,mes,cod_cuenta,saldo_inicial,total_debitos,total_creditos" _
              & ",saldo_final) values(" & lngAnio & "," & iMes & ",'" & Trim(rs!fnd_cuenta) & "'," _
              & "0,"
      If rs!fnd_DEBEHABER = "D" Then
          strSQL = strSQL & rs!Movimiento & ",0,0)"
      Else
          strSQL = strSQL & "0," & rs!Movimiento & ",0)"
      End If
              
   End If 'Existe cuenta

  Call ConectionExecute(strSQL)

  prgBar.Value = prgBar.Value + 1
  rs.MoveNext

Loop
rs.Close


'Procesar la Cola de Asientos
lbl.Caption = "Procesando Resumen de Movimientos de Cuentas (II)..."
lbl.Refresh

strSQL = "select A.fnd_cuenta,A.fnd_debehaber,isnull(sum(A.fnd_monto),0) as Movimiento" _
       & " from fnd_asientos_cola A" _
       & " where datepart(yyyy, A.fnd_fecha) = " & lngAnio _
       & " and datepart(mm, A.fnd_fecha) = " & iMes & " group by A.fnd_cuenta,A.fnd_debehaber"

Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
    
Do While Not rs.EOF
     
   If fxExisteCuenta(iMes, lngAnio, rs!fnd_cuenta) Then
      'Existe
      If rs!fnd_DEBEHABER = "D" Then
        strSQL = "update fnd_per_cuentas set total_debitos = total_debitos + " & rs!Movimiento _
               & " where anio = " & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
               & Trim(rs!fnd_cuenta) & "'"
      Else
        strSQL = "update fnd_per_cuentas set total_creditos = total_creditos + " & rs!Movimiento _
               & " where anio = " & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
               & Trim(rs!fnd_cuenta) & "'"
      End If
   
   Else 'No existe la cuenta
       strSQL = "insert fnd_per_cuentas(anio,mes,cod_cuenta,saldo_inicial,total_debitos,total_creditos" _
              & ",saldo_final) values(" & lngAnio & "," & iMes & ",'" & Trim(rs!fnd_cuenta) & "'," _
              & "0,"
      If rs!fnd_DEBEHABER = "D" Then
          strSQL = strSQL & rs!Movimiento & ",0,0)"
      Else
          strSQL = strSQL & "0," & rs!Movimiento & ",0)"
      End If
              
   End If 'Existe cuenta

  Call ConectionExecute(strSQL)
  prgBar.Value = prgBar.Value + 1
  rs.MoveNext
Loop
rs.Close


'Revisando la clasificacion de las cuentas

lbl.Caption = "Procesando Resumen de Movimientos de Cuentas (II)..."
lbl.Refresh

strSQL = "select A.cod_cuenta,T.clasificacion" _
       & " from fnd_per_cuentas A inner join cuentas C on A.cod_cuenta = C.cod_cuenta" _
       & " inner join Tipos_cuenta T on C.tipo_cuenta = T.tipo_cuenta" _
       & " where A.anio = " & lngAnio & " and Mes = " & iMes _
       & " and A.clasificacion is null"

Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
    
Do While Not rs.EOF
  strSQL = "update fnd_per_cuentas set clasificacion = '" & rs!Clasificacion _
         & "' where anio = " & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
         & Trim(rs!cod_cuenta) & "'"
  Call ConectionExecute(strSQL)
  prgBar.Value = prgBar.Value + 1
  rs.MoveNext
Loop
rs.Close


'Cerrando el Periodo e Iniciando el nuevo


lbl.Caption = "Cerrando Periodo e Iniciando Nuevo Periodo (I)..."
lbl.Refresh

strSQL = "select *  from fnd_per_cuentas" _
       & " where anio = " & lngAnio & " and Mes = " & iMes

Call OpenRecordSet(rs, strSQL)
prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
    
Do While Not rs.EOF
  Select Case rs!Clasificacion
     Case "G", "A", "O", "V"
        strSQL = "update fnd_per_cuentas set saldo_final = Saldo_inicial + Total_debitos - Total_creditos" _
               & " where anio = " & lngAnio & " and Mes = " & iMes & " and cod_cuenta = '" _
               & rs!cod_cuenta & "'"
     Case "I", "C", "P"
        strSQL = "update fnd_per_cuentas set saldo_final = Saldo_inicial - Total_debitos + Total_creditos" _
               & " where anio = " & lngAnio & " and Mes = " & iMes & " and cod_cuenta = '" _
               & rs!cod_cuenta & "'"
  End Select
  Call ConectionExecute(strSQL)
  prgBar.Value = prgBar.Value + 1
  rs.MoveNext
Loop
rs.Close

lbl.Caption = "Cerrando Periodo e Iniciando Nuevo Periodo (II)..."
lbl.Refresh
'Inicializa el proximo periodo
If iMes = 12 Then
   iMesN = 1
   lngAnioN = lngAnio + 1
Else
   iMesN = iMes + 1
   lngAnioN = lngAnio
End If


strSQL = "insert into fnd_per_cuentas(cod_cuenta,anio,mes,saldo_inicial,total_creditos,total_debitos,Saldo_final,clasificacion) " _
       & "select cod_cuenta," & lngAnioN & "," & iMesN & ",saldo_final,0,0,0,clasificacion from fnd_per_cuentas where anio = " _
       & lngAnio & " and mes = " & iMes
Call ConectionExecute(strSQL)

'''''' ****** Se uso para corregir *******
'''''strSQL = "select * from fnd_per_cuentas where anio = " & lngAnio & " and mes = " & iMes
'''''Call OpenRecordSet(rs, strSQL)
'''''Do While Not rs.EOF
'''''    strSQL = "update fnd_per_cuentas set saldo_inicial = " & rs!saldo_final _
'''''           & " where Anio = " & lngAnioN & " and mes = " & iMesN _
'''''           & " and cod_cuenta = '" & rs!cod_cuenta & "'"
'''''    Call ConectionExecute(strSQL)
'''''rs.MoveNext
'''''Loop
'''''rs.Close

End Sub


Private Sub cmdCierre_Click()
Dim strSQL As String, iMes As Integer, lngAnio As Long
Dim rs As New ADODB.Recordset, iRespuesta As Integer

'PASOS
'1. Guardar El Estado Actual de los Creditos
'   saldo_inicial,saldo,proceso,opex,id_solicitud,codigo
'2. Actualizar Total_debitos y Total_creditos con los
'   movimientos del mes
'3. Establecer Nuevo Corte de Saldos.
'4. Insertar en Historicos, el periodo procesado.
'5. Crear Referencia Contable (Metodo Contable)


iRespuesta = MsgBox("Esta seguro que desea establecer cierre del Mes, se le recuerda" _
                   & " que tiene que ser el ultimo día del mes, cuando ya no se procese información", vbYesNo)
If iRespuesta = vbNo Then Exit Sub


lbl.Alignment = 0

lbl.Caption = "Cargando parámetros y actualizando últimos movimientos..."
lbl.Refresh


iMes = Month(fxFechaServidor)
lngAnio = Year(fxFechaServidor)

strSQL = "select isnull(count(*),0) as Existe from fnd_per_historico where anio = " _
       & lngAnio & " and mes = " & iMes
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  rs.Close
  MsgBox "El Procedimiento de Cierre ya se Corrio para este mes...", vbExclamation
  Exit Sub
End If
rs.Close


Me.MousePointer = vbHourglass

'Borrando Información Anterior (no deberia de existir)
strSQL = "delete FND_per_cerrados where anio = " & lngAnio & " and mes = " & iMes
Call ConectionExecute(strSQL)


lbl.Caption = "Copiando Estado Actual..."
lbl.Refresh


'Guarda los datos de los aportes en Historial de Aportes
strSQL = "insert into FND_per_cerrados(anio,mes,cod_operadora,cod_plan,cod_contrato,aportes,rendimientos,estado)" _
       & " (select " & lngAnio & "," & iMes & ",cod_operadora,cod_plan,cod_contrato,aportes,rendimiento,estado" _
       & " from fnd_contratos)"
Call ConectionExecute(strSQL)

lbl.Caption = "Registrando ingreso a historicos"
lbl.Refresh

strSQL = "insert into fnd_per_historico(anio,mes) values(" & lngAnio _
       & "," & iMes & ")"
Call ConectionExecute(strSQL)


'Nuevos Proceso de Saldos x Cuentas

'Call sbMovCuentas

lbl.Caption = "Cierre Concluido Satisfactoriamente...."
prgBar.Value = 1
prgBar.Max = 1000000

Me.MousePointer = vbDefault

Exit Sub


vError:
  lbl.Caption = "Error...."
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Function fxTmpContrato(vCedula As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_contrato from fnd_contratos where cedula = '" _
       & vCedula & "' and cod_plan = 'ANAV 05'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxTmpContrato = 0
Else
  fxTmpContrato = rs!COD_CONTRATO
End If
rs.Close

End Function

Private Function fxTmpIDDetalle(vContrato As Long, vMonto As Currency, vFecha As Date) As Long
'Dim strSQL As String, rs As New ADODB.Recordset
'
'strSQL = "select cod_fnd_Detalle from fnd_contratos_detalle where cod_plan = 'ANAV 05' and monto = " _
'       & vMonto & " and fecha = '" & Format(vFecha, "yyyy/mm/dd") & "' and cod_contrato = 1635 and tcon = 1"
'Call OpenRecordSet(rs, strSQL)
'If rs.EOF And rs.BOF Then
'  fxTmpIDDetalle = 0
'Else
'  fxTmpIDDetalle = rs!cod_fnd_Detalle
'End If
'rs.Close

End Function


Private Sub Command1_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vContrato As Long, vDetalle As Long

strSQL = "select C.cedula,D.* " _
       & " from creditos_Dt D inner join reg_creditos C on D.id_solicitud = C.id_solicitud" _
       & " where D.codigo = 'ANAV' and D.fechas between '2004/12/01' and '2005/11/30'" _
       & " and D.Tcon = 1"
Call OpenRecordSet(rs, strSQL)

prgBar.Value = 1
prgBar.Max = rs.RecordCount + 1
Do While Not rs.EOF
 vContrato = fxTmpContrato(rs!Cedula)
 If vContrato > 0 Then
   vDetalle = fxTmpIDDetalle(vContrato, rs!Abono, rs!fechas)
   If vDetalle > 0 Then
     strSQL = "update fnd_Contratos_Detalle set cod_contrato = " & vContrato _
            & " where cod_fnd_detalle = " & vDetalle
     Call ConectionExecute(strSQL)
   End If
   
 End If
 
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

MsgBox "Fin"



End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

lbl.Caption = "Este proceso realiza el cierre del mes, el cual crea una copia para uso historico" _
            & " por cortes mensuales" & vbCrLf & vbCrLf & " POR ESTE MOTIVO ESTE PROCESO SOLO DEBE DE SER" _
            & " REALIZADO EL ULTIMO DIA DEL MES (UNA SOLA VEZ)"

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
