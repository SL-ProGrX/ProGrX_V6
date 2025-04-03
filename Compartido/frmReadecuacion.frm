VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReadecuacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Readecuaciones 03/2001"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbl 
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
End
Attribute VB_Name = "frmReadecuacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub sbotro()
Dim rs As New ADODB.Recordset
Dim Moro As New ADODB.Recordset
Dim i As Integer
Dim T As Integer

strSQL = "select count(*) as cuotas,M.id_solicitud from "
strSQL = strSQL & "morosidad M inner join reg_creditos R "
strSQL = strSQL & "on M.id_solicitud=R.id_solicitud inner join temporalced T "
strSQL = strSQL & "on R.cedula=T.Cedula inner join Catalogo C "
strSQL = strSQL & "on R.codigo=C.codigo inner join Socios S "
strSQL = strSQL & "on S.Cedula=R.Cedula "
strSQL = strSQL & "where M.estado='A' and T.prestamos > 0 "
strSQL = strSQL & "and R.Proceso='N' and M.Fechap=200103 and "
strSQL = strSQL & "C.Retencion = 'N' And C.Poliza = 'N' And S.EstadoActual <> 'N' "
strSQL = strSQL & "group by M.id_solicitud"

With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    T = .RecordCount
    prgBar.Max = T
    Do While .EOF = False
       Moro.Open "Select * From Vista_Morosidad Where Id_solicitud=" & !id_solicitud, glogon.Conection, adOpenStatic
         If Moro.EOF = False And Moro!Cuota <= 2 Then
            Call EjecuteOperacion(!id_solicitud, Moro!IntC, Moro!Intm)
         End If
       Moro.Close
       .MoveNext
       i = i + 1
       prgBar.Value = i
       lbl = i & " DE " & T
       Me.Refresh
    Loop
 .Close
End With

MsgBox "FIN"


End Sub
Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As String, vPlazo As Integer, rs2 As New ADODB.Recordset
Dim vFechaProceso As Long, vFechaInicio As String, vFechaCorte As String
Dim vCuota As Currency

frmLogon.Show vbModal

strSQL = "select id_solicitud,observacion,interesv,montoapr  From reg_creditos " _
       & " Where observacion Like 'READECUACION DE LA OPERACION x INCON 200103 CED%'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  vOperacion = Mid(rs!observacion, 50, 8)
  
  strSQL = "select plazo,prideduc,fecult from reg_creditos where id_solicitud = " & vOperacion
  rs2.Open strSQL, glogon.Conection, adOpenStatic
   If rs2!FECULT < 200103 Then
       vFechaProceso = 200103
   Else
       vFechaProceso = rs2!FECULT
   End If
   vFechaInicio = "01/" & Mid(CStr(rs2!prideduc), 5, 2) & "/" & Mid(CStr(rs2!prideduc), 1, 4)
   vFechaCorte = "01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4)
   
   vPlazo = rs2!Plazo - (DateDiff("m", vFechaInicio, vFechaCorte))
   vPlazo = vPlazo + 1
   
   If vPlazo <= 0 Then vPlazo = 1
   
   vCuota = CCur(fxCalcula_Cuota(rs!montoapr, vPlazo, rs!interesv))
   strSQL = "update reg_creditos set cuota =" & vCuota & ",plazo = " & vPlazo _
          & " where id_solicitud = " & rs!id_solicitud
   glogon.Conection.Execute strSQL
  
  rs2.Close
  
  rs.MoveNext
Loop
rs.Close

MsgBox "f"

End Sub


Function fxUltimaOperacion(strCedula As String) As Long
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select coalesce(max(id_solicitud),0) as operacion from reg_creditos " _
       & "where cedula ='" & Trim(strCedula) & "'"

rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
       
fxUltimaOperacion = rsX!Operacion
rsX.Close
  
End Function

Sub AsientoReadecuacion(vOpex As Integer, curIntc As Currency, curIntm As Currency _
    , curSaldo As Currency, vCodigo As String, vOperacion As Long)
Dim rs As New ADODB.Recordset, strSQL As String, curMonto As Currency
Dim vFecha As Date


vFecha = fxFechaServidor

curMonto = curIntc + curIntm + curSaldo

If vOpex = 0 Then
  strSQL = "select ctanintc as ctaintc,ctanintm as ctaintm, ctanamort as ctaAmortiza "
Else 'cuentas exsocios
  strSQL = "select ctaointc as ctaintc,ctaointm as ctaintm, ctaoamort as ctaAmortiza "
End If

strSQL = strSQL & "from catalogo where codigo = '" & vCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
    & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
    & "','RA" & vOperacion & "','" & rs!ctaamortiza & "'," & curMonto & ",'D','" _
    & Format(vFecha, "mm/dd/yyyy") & "','P')"
glogon.Conection.Execute strSQL

strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
    & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
    & "','RA" & vOperacion & "','" & rs!ctaamortiza & "'," & curSaldo & ",'H','" _
    & Format(vFecha, "mm/dd/yyyy") & "','P')"
glogon.Conection.Execute strSQL


If curIntc > 0 Then
strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
    & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
    & "','RA" & vOperacion & "','" & rs!ctaintc & "'," & curIntc & ",'H','" _
    & Format(vFecha, "mm/dd/yyyy") & "','P')"
glogon.Conection.Execute strSQL
End If


If curIntm > 0 Then
strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
    & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
    & "','RA" & vOperacion & "','" & rs!ctaintm & "'," & curIntm & ",'H','" _
    & Format(vFecha, "mm/dd/yyyy") & "','P')"
glogon.Conection.Execute strSQL
End If

rs.Close

End Sub


Function fxPrimerDeduccion() As String
Dim Mes As Long, Anio As Long, Fecha As String

Mes = Val(Mid(CStr(200103), 5, 2))
Anio = Val(Mid(CStr(200103), 1, 4))

If Mes = 12 Then
 Anio = Anio + 1
 Mes = 1
Else
 Mes = Mes + 1
End If

If Mes < 10 Then
 Fecha = Trim(CStr(Anio)) + "0" + Trim(CStr(Mes))
Else
 Fecha = Trim(CStr(Anio)) + Trim(CStr(Mes))
End If

fxPrimerDeduccion = Fecha

End Function

Function fxCalcula_Cuota(Monto As Double, Plazo As Integer, Interes As Integer)
Dim curCuota As Double, curInteresMensual As Double, i As Integer
Dim curfactor As Double, curTotalInteresmensual As Double, curTotalamortizacion As Double
Dim curvalorfuturo As Double

On Error Resume Next

curInteresMensual = Interes / (12 * 100)
curfactor = 1

For i = 1 To Plazo
  curfactor = curfactor * (1 + curInteresMensual)
Next i
 curCuota = Monto * ((curInteresMensual * curfactor) / (curfactor - 1))
 curCuota = IIf((Interes = 0), (Monto / Plazo), curCuota)
 fxCalcula_Cuota = Format(curCuota, "###,###,###,##0.00")
End Function



Sub EjecuteOperacion(vOperacion As Long, vIntC As Currency, vIntM As Currency)
Dim rsEO As New ADODB.Recordset, strSQL As String
Dim strCedula As String, strOB As String, vOpex As Integer
Dim lngUltimaOperacion As Long, rs As New ADODB.Recordset
Dim vFecha As Date, rs2 As New ADODB.Recordset
Dim vSaldo As Currency, vCodigo As String
Dim vPlazo As Integer, vFechaProceso As Long


glogon.Conection.BeginTrans

On Error GoTo vError

vFecha = fxFechaServidor
  
 strSQL = "select * from reg_creditos where id_solicitud = " & vOperacion
 rs2.Open strSQL, glogon.Conection, adOpenStatic
   
 'Guarda el datos antes de los movimientos
 vSaldo = rs2!saldo
 vCodigo = rs2!codigo
 strCedula = rs2!Cedula
   
   
'Cancelar morosidad sin abonos
 strSQL = "update morosidad set estado = 'C'," _
        & "abintc = intc,abintm = intm,abamortiza = amortiza,tcon =4,ncon=8889" _
        & ",fecult = '" & Format(vFecha, "mm/dd/yyyy") & "' " _
        & "where estado = 'A' and id_solicitud = " & vOperacion
 glogon.Conection.Execute strSQL

 strOB = "READECUACION DE LA OPERACION x INCON 200103 CED : " & vOperacion

'Cancelo la operacion actual
 strSQL = "update reg_creditos set saldo = 0, amortiza = montoapr,saldo_mes = 0," _
        & "estado = 'C',FECHA_ENVIAPROCESO = '" & Format(vFecha, "mm/dd/yyyy") _
        & "',OBSERVACION_PROCESO='Se readecuo la deuda' " _
        & "where id_solicitud = " & vOperacion
 glogon.Conection.Execute strSQL


rs.CursorLocation = adUseServer
rs.Open "select coalesce(sum(abamortiza),0) as Amortiza from morosidad where tcon = 4" _
        & " and ncon = 8889 and id_solicitud = " & vOperacion, glogon.Conection, adOpenStatic
'Insertar Registro de Detalle

strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
       & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & rs2!codigo & "'," _
       & vOperacion & ",0," & rs2!saldo - rs!Amortiza _
       & ",0," & rs2!saldo - rs!Amortiza & ",'" & Format(vFecha, "mm/dd/yyyy") _
       & "',200103,4,8889,'A','G')"
glogon.Conection.Execute strSQL
rs.Close

rs2.Close
'Abrir nueva operacion

 rsEO.Source = "select * from reg_creditos where id_solicitud = " & vOperacion
 rsEO.Open , glogon.Conection, adOpenStatic
  
 vOpex = IIf(IsNull(rsEO!opex), 0, rsEO!opex)
 
 '************************ REVISAR ESTA PARTE CON EL COMITE DE INFORMATICA
 
 With rsEO
'    Mes = Val(Mid(CStr(200103), 5, 2))
'    Anio = Val(Mid(CStr(200103), 1, 4))
   
   If !FECULT < 200103 Then
       vFechaProceso = 200103
   Else
       vFechaProceso = !FECULT
   End If
   vPlazo = 2 + DateDiff("m", "01/" & Mid(CStr(!prideduc), 5, 2) & "/" & Mid(CStr(!prideduc), 1, 4) _
            , "01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4))
 
 
    strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,estadosol,fechasol,fechares,"
    strSQL = strSQL & "plazo,int,montoapr,prideduc,fechaforp,fechaforf,acta,saldo,amortiza,interesc,"
    strSQL = strSQL & "cuota,estado,opex,proceso,userrec,userres,userfor,garantia,observacion,"
    strSQL = strSQL & "firma_deudor,monto_girado,interesv,tesoreria,usertesoreria,primer_cuota,"
    strSQL = strSQL & "tdocumento,ndocumento,pagare,fecha_calculo_int,premio,"
    strSQL = strSQL & "cuotas_planilla,cuotas_directas,cuotas_anuladas,FECULT) values("
 
 
   strSQL = strSQL & "'" & !codigo & "'," & !id_comite & ",'" & !Cedula & "'," & (vSaldo + vIntC + vIntM) & ",'F','" & Format(vFecha, "mm/dd/yyyy") & "','" & Format(vFecha, "mm/dd/yyyy") & "',"
   strSQL = strSQL & vPlazo & "," & !Int & "," & (vSaldo + vIntC + vIntM) & ",200104,'" & Format(vFecha, "mm/dd/yyyy") & "','" & Format(vFecha, "mm/dd/yyyy") & "',"
   strSQL = strSQL & IIf(IsNull(!acta), 0, !acta) & "," & (vSaldo + vIntC + vIntM) & ",0,0,"
   strSQL = strSQL & CCur(fxCalcula_Cuota((vSaldo + vIntC + vIntM), vPlazo, !interesv)) & ",'A'," & !opex & ",'N','" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "','" & !garantia & "','" & strOB & "',"
   strSQL = strSQL & "1,0," & !interesv & ",'" & Format(vFecha, "mm/dd/yyyy") & "','" & glogon.Usuario & "','N','"
   strSQL = strSQL & !Tdocumento & "','" & !ndocumento & "'," & IIf(IsNull(!pagare), 0, !pagare) & ",'" & Format(vFecha, "mm/dd/yyyy") & "'," & IIf(IsNull(!premio), 0, !premio) & ","
   strSQL = strSQL & "0,0,0," & vFechaProceso & ")"
 End With
 
 '***************************************************************************

 glogon.Conection.Execute strSQL

 rsEO.Close
 
'Recuperar la nueva operacion

   lngUltimaOperacion = fxUltimaOperacion(strCedula)

  'Hereda Fiadores Operacion Anterior
  With rsEO
   .CursorLocation = adUseServer
   .Source = "select * from fiadores where id_solicitud = " & vOperacion
   .Open , glogon.Conection, adOpenStatic
   Do While Not .EOF
    strSQL = "insert fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado) values(" _
           & lngUltimaOperacion & ",'" & !codigo & "','" & !cedulaf _
           & "','" & !Nombre & "','" & !firma & "','" & !estado & "')"
    glogon.Conection.Execute strSQL
    .MoveNext
   Loop
   .Close
  End With

  
  Call AsientoReadecuacion(vOpex, vIntC, vIntM, vSaldo, vCodigo, vOperacion)
  


  glogon.Conection.CommitTrans
  
Exit Sub

vError:
 glogon.Conection.RollbackTrans
 MsgBox Err.Description, vbCritical

End Sub



