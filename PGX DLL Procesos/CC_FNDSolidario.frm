VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCC_FNDSolidario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondo Solidario"
   ClientHeight    =   2004
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7980
   Icon            =   "CC_FNDSolidario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2004
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   216
      Left            =   0
      TabIndex        =   2
      Top             =   1788
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   381
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Picture         =   "CC_FNDSolidario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1440
      X2              =   8040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "ESTE PROCESO DE DEBE DE CORRER DESPUES DE LA APLICACION MENSUAL DE CREDITOS Y ANTES DE ENVIAR LAS CUOTAS AL COBRO."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "frmCC_FNDSolidario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxFondoSolidario(curMonto As Currency, Optional curMontoBase As Currency = 150) As Currency
fxFondoSolidario = (curMonto / 1000000) * curMontoBase
End Function


Private Sub sbFondoBeneficioSocial()
Dim strSQL As String, i As Integer
Dim vCodigo As String, vMonto As Currency
Dim vFecha As Date


i = MsgBox("Esta seguro que desea realizar el proceso del Fondo de Beneficio Social, se le recuerda" _
                   & " que tiene que ser despues de la aplicacion mensual (Abonos por Planilla)", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


vCodigo = "FBEN"
vMonto = 800
vFecha = fxFechaServidor

lbl.Caption = "Excluyendo Cuotas de Ex-Socios..."
lbl.Refresh
 
strSQL = "update reg_creditos set estado = 'C',saldo = 0,cuota = 0 where Estado = 'A' and codigo = '" _
       & vCodigo & "' and cedula in(select cedula from socios where estadoactual <> 'S')"
Call ConectionExecute(strSQL)
 
'ACTUALIZA MONTO CASOS ACTUALES
strSQL = "update reg_creditos set montoapr = " & vMonto _
       & ",cuota = " & vMonto & ",saldo = " & vMonto _
       & " where Estado = 'A' and codigo = '" & vCodigo _
       & "' and cedula in(select cedula from socios where estadoactual = 'S')"
Call ConectionExecute(strSQL)
  
 
lbl.Caption = "Procesando Casos Nuevos..."
lbl.Refresh
 
strSQL = "insert into reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
       & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
       & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
       & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
       & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,documento_referido)" _
       & " (select '" & UCase(vCodigo) & "',6,cedula," & vMonto & "," & vMonto & ",0," & vMonto & ",0,0," _
       & vMonto & "," & vMonto & ",0,0,999,'" & glogon.Usuario & "','" & glogon.Usuario _
       & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','N'" _
       & ",'N','OT','',0,1,0,'Proceso Automatico Cuota Mantenimiento CR','A'," & fxFechaProcesoSiguiente(GLOBALES.glngFechaCR) _
       & "," & fxFechaProcesoAnterior(GLOBALES.glngFechaCR) & ",'F','AUTOMATICO'" _
       & " from socios where estadoactual = 'S' and cedula not in(select cedula from reg_creditos" _
       & " where estado = 'A' and codigo = '" & vCodigo & "'))"
Call ConectionExecute(strSQL)
 

'Cancela Casos en Congelamiento Activado

'strSQL = "update reg_creditos set estado = 'C' where estado = 'A' and codigo = '" & vCodigo _
'       & "' and cedula in(select cedula From afi_congelar where estado = 'A'  and fecha_finaliza >= dbo.MyGetdate()" _
'       & "  and per_cobro_cuotaCr = 0)"
'Call ConectionExecute(strSQL)

'Cancela Casos sin deduccion de Aportes por mas de 2 Meses

strSQL = "update reg_creditos set estado = 'C',saldo = 0 where estado = 'A' and codigo = '" & vCodigo _
       & "' and cedula in(select A.cedula From ahorro_consolidado A inner join socios S on A.cedula = S.cedula" _
       & " where S.estadoactual = 'S'  and datediff( m, A.fecAporte, dbo.MyGetdate()) > 2 )"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
lbl.Caption = "Fondo de Beneficio Socual Actualizado Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbFondoSolidario()
Dim strSQL As String, i As Integer, vCodigo As String


i = MsgBox("Esta seguro que desea realizar el calculo del fondo solidario, se le recuerda" _
                   & " que tiene que ser despues de la aplicacion mensual (Abonos por Planilla)", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información..."
lbl.Refresh

vCodigo = "FNDS"

'NUEVO : Modificado el 23/08/2006
'        Aplica con parametro de Cobertura  del catalogo de creditos
' Actualizado el 28/07/2009 x instituciones

glogon.Conection.CommandTimeout = 5000
 

'Cancela Operaciones de FNDS que no tienen deudas con cobertura
       
strSQL = "update R set Estado = 'C'" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and R.codigo = '" & vCodigo & "' and R.estado = 'A' and R.cedula not in(Select Reg.Cedula" _
       & " from reg_creditos Reg inner join catalogo Cat on Reg.codigo = Cat.codigo" _
       & " where Cat.retencion = 'N' and Cat.poliza = 'N' and Cat.cobertura = 1 and Reg.garantia not in('H')" _
       & " and Reg.saldo > 0 and Reg.estado = 'A' and Reg.proceso <> 'J'" _
       & " group by Reg.cedula)"
Call ConectionExecute(strSQL)


lbl.Caption = "Inicializando Cuotas Activas del Fondo..."
lbl.Refresh


'Inicializa Cuotas
strSQL = "update R set cuota = 0, saldo = 0,montoapr = 0,saldo_mes = 0" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado = 'A' and R.codigo = '" & vCodigo & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
Call ConectionExecute(strSQL)

Call sbFondoSolidarioPaso1
Call sbFondoSolidarioPaso2
Call sbFondoSolidarioPaso3


'Cancela Casos en Congelamiento Activado para Fondo Solidario
strSQL = "update reg_creditos set estado = 'C'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado = 'A' and R.codigo = '" & vCodigo _
       & "' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "  and R.cedula in(select cedula From afi_congelar where estado = 'A'  and fecha_finaliza >= dbo.MyGetdate()" _
       & "  and per_cobro_fndSol = 0)"
Call ConectionExecute(strSQL)


prgBar.Value = 1

Me.MousePointer = vbDefault
lbl.Caption = "Fondo Solidario Actualizado Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub cmdAplicar_Click()


If GLOBALES.SysASEVersion Then
   Call sbFondoSolidario
Else
   'ASEASECCSS / No Aplica (Proc. Add. Planilla)
   Call sbFondoBeneficioSocial
End If

End Sub


Private Sub sbFondoSolidarioPaso1()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset, iRespuesta As Integer
Dim vCodigo As String, vFecha As Date
Dim vFND As Currency, vFechaProX As Long

'Aplica Fondo Para Garantias en el Ahorro y Sobre Saldos


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información..."
lbl.Refresh

vCodigo = "FNDS"
vFecha = fxFechaServidor
vFechaProX = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

'Actualizando FND
strSQL = "select R.cedula,sum(R.montoapr) as Monto" _
       & " from reg_Creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " where C.retencion = 'N' and C.poliza = 'N'  and C.cobertura = 1 and R.garantia in('A','N')" _
       & " and R.saldo > 0 and R.estado = 'A' and R.proceso <> 'J' and R.fechaforp < '2004/06/01'" _
       & " group by R.cedula"
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1


Do While Not rs.EOF
 
 lbl.Caption = "Actualizando FNDS..." & vbCrLf & vbCrLf & "Procesando " & prgBar.Value & " de " & prgBar.Max
 lbl.Refresh
 
 vFND = fxFondoSolidario(rs!Monto, 150)
 
 strSQL = "Select id_solicitud,cuota from reg_creditos where codigo = '" & vCodigo & "' and cedula = '" _
        & rs!Cedula & "' and estado = 'A'"
 rsTmp.Open strSQL, glogon.Conection, adOpenStatic
 If rsTmp.EOF And rsTmp.BOF Then
   'Insertar nueva cuota
     strSQL = "insert reg_creditos(id_comite,codigo,cedula,montosol,montoapr,plazo,int,interesv" _
          & ",saldo,interesc,amortiza,cuota,prideduc,fecult,estadosol,estado,fechasol,fechares,fechaforp" _
          & ",fechaforf,observacion,garantia,tdocumento,ndocumento,tesoreria,userrec,userfor,userres)" _
          & "values(1,'" & vCodigo & "','" & Trim(rs!Cedula) & "'," & vFND & "," & vFND & ",999,0,0," & vFND _
          & ",0,0," & vFND & "," & vFechaProX & "," & GLOBALES.glngFechaCR & ",'F','A','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") _
          & "','FONDO SOLIDARIO CREADO EL " & vFecha & "','S','OT','','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "')"
      Call ConectionExecute(strSQL)

 Else
   If Abs(vFND - rsTmp!Cuota) > 1 Then
     'Actualizar la cuota
     strSQL = "update reg_creditos set cuota = " & vFND & ", saldo = " & vFND & ", montoapr = " & vFND _
            & ", saldo_mes = " & vFND _
            & " where id_solicitud = " & rsTmp!Id_solicitud
     Call ConectionExecute(strSQL)
   End If
 End If
 rsTmp.Close
 
 prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close


prgBar.Value = 1

Me.MousePointer = vbDefault
lbl.Caption = "Fondo Solidario Paso 1 Actualizado Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbFondoSolidarioPaso2()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset, iRespuesta As Integer
Dim vCodigo As String, vFecha As Date
Dim vFND As Currency, vFechaProX As Long

'Aplica Fondo Para Garantias Diff del Ahorro y Sobre Monto Aprobado

Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información Paso 2..."
lbl.Refresh

vCodigo = "FNDS"
vFecha = fxFechaServidor
vFechaProX = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

'Actualizando FND
strSQL = "select R.cedula,sum(R.montoapr) as Monto" _
       & " from reg_Creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " where C.retencion = 'N' and C.poliza = 'N' and C.cobertura = 1 and R.garantia in('F','X')" _
       & " and R.saldo > 0 and R.estado = 'A' and R.proceso <> 'J' and R.fechaforp < '2004/06/01'" _
       & " group by R.cedula"
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1


Do While Not rs.EOF
 
 lbl.Caption = "Actualizando FNDS..." & vbCrLf & vbCrLf & "Procesando " & prgBar.Value & " de " & prgBar.Max
 lbl.Refresh
 
 vFND = fxFondoSolidario(rs!Monto, 300)
 
 strSQL = "Select id_solicitud,cuota from reg_creditos where codigo = '" & vCodigo & "' and cedula = '" _
        & rs!Cedula & "' and estado = 'A'"
 rsTmp.Open strSQL, glogon.Conection, adOpenStatic
 If rsTmp.EOF And rsTmp.BOF Then
   'Insertar nueva cuota
     strSQL = "insert reg_creditos(id_comite,codigo,cedula,montosol,montoapr,plazo,int,interesv" _
          & ",saldo,interesc,amortiza,cuota,prideduc,fecult,estadosol,estado,fechasol,fechares,fechaforp" _
          & ",fechaforf,observacion,garantia,tdocumento,ndocumento,tesoreria,userrec,userfor,userres)" _
          & "values(1,'" & vCodigo & "','" & Trim(rs!Cedula) & "'," & vFND & "," & vFND & ",999,0,0," & vFND _
          & ",0,0," & vFND & "," & vFechaProX & "," & GLOBALES.glngFechaCR & ",'F','A','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") _
          & "','FONDO SOLIDARIO CREADO EL " & vFecha & "','Z','OT','','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "')"
      Call ConectionExecute(strSQL)

 Else
   If Abs(vFND - rsTmp!Cuota) > 1 Then
     'Actualizar la cuota
     strSQL = "update reg_creditos set cuota = cuota + " & vFND _
            & ", saldo = saldo + " & vFND & ",saldo_mes = saldo_mes + " & vFND _
            & ", montoapr = montoapr + " & vFND _
            & " where id_solicitud = " & rsTmp!Id_solicitud
     Call ConectionExecute(strSQL)
   End If
 End If
 rsTmp.Close
 
 prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close


prgBar.Value = 1

Me.MousePointer = vbDefault
lbl.Caption = "Fondo Solidario Actualizado Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbFondoSolidarioPaso3()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset, iRespuesta As Integer
Dim vCodigo As String, vFecha As Date
Dim vFND As Currency, vFechaProX As Long

'Aplica Fondo Para Garantias Diff del Ahorro y Sobre Monto Aprobado

Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información Paso 2..."
lbl.Refresh

vCodigo = "FNDS"
vFecha = fxFechaServidor
vFechaProX = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

'Actualizando FND
strSQL = "select R.cedula,sum(R.montoapr) as Monto" _
       & " from reg_Creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " where C.retencion = 'N' and C.poliza = 'N' and C.cobertura = 1 and R.garantia not in('H')" _
       & " and R.saldo > 0 and R.estado = 'A' and R.proceso <> 'J' and R.fechaforp >= '2004/06/01'" _
       & " group by R.cedula"
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1


Do While Not rs.EOF
 
 lbl.Caption = "Actualizando FNDS..." & vbCrLf & vbCrLf & "Procesando " & prgBar.Value & " de " & prgBar.Max
 lbl.Refresh
 
 vFND = fxFondoSolidario(rs!Monto, 300)
 
 strSQL = "Select id_solicitud,cuota from reg_creditos where codigo = '" & vCodigo & "' and cedula = '" _
        & rs!Cedula & "' and estado = 'A'"
 rsTmp.Open strSQL, glogon.Conection, adOpenStatic
 If rsTmp.EOF And rsTmp.BOF Then
   'Insertar nueva cuota
     strSQL = "insert reg_creditos(id_comite,codigo,cedula,montosol,montoapr,plazo,int,interesv" _
          & ",saldo,interesc,amortiza,cuota,prideduc,fecult,estadosol,estado,fechasol,fechares,fechaforp" _
          & ",fechaforf,observacion,garantia,tdocumento,ndocumento,tesoreria,userrec,userfor,userres)" _
          & "values(1,'" & vCodigo & "','" & Trim(rs!Cedula) & "'," & vFND & "," & vFND & ",999,0,0," & vFND _
          & ",0,0," & vFND & "," & vFechaProX & "," & GLOBALES.glngFechaCR & ",'F','A','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") _
          & "','FONDO SOLIDARIO CREADO EL " & vFecha & "','Z','OT','','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "')"
      Call ConectionExecute(strSQL)

 Else
   If Abs(vFND - rsTmp!Cuota) > 1 Then
     'Actualizar la cuota
     strSQL = "update reg_creditos set cuota = cuota + " & vFND _
            & ", saldo = saldo + " & vFND _
            & ", montoapr = montoapr + " & vFND _
            & " where id_solicitud = " & rsTmp!Id_solicitud
     Call ConectionExecute(strSQL)
   End If
 End If
 rsTmp.Close
 
 prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close


prgBar.Value = 1

Me.MousePointer = vbDefault
lbl.Caption = "Fondo Solidario Actualizado Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub




Private Sub Form_Load()
Dim strSQL As String
 
 vModulo = 10
 
   Me.Caption = "Fondo Solidario"

 If GLOBALES.SysASEVersion Then
   Me.Caption = "Fondo Solidario"
 Else
   Me.Caption = "Fondo de Beneficio Social"
 End If

 
'Solo Activa las Instituciones CCSS y OPC
strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX " _
       & " from instituciones where activa = 1 and cod_institucion in(1,2) order by descripcion"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub
