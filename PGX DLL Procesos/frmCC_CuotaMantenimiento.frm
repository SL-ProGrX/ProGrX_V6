VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCC_CuotaMantenimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de la Cuota de Mantenimiento Centro de Recreo"
   ClientHeight    =   2016
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8160
   Icon            =   "frmCC_CuotaMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2016
   ScaleWidth      =   8160
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
      Picture         =   "frmCC_CuotaMantenimiento.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   204
      Left            =   0
      TabIndex        =   0
      Top             =   1812
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   360
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
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
      Caption         =   $"frmCC_CuotaMantenimiento.frx":0572
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
      TabIndex        =   2
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "frmCC_CuotaMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxFondoSolidario(curMonto As Currency, Optional curMontoBase As Currency = 150) As Currency
fxFondoSolidario = (curMonto / 1000000) * curMontoBase
End Function


Private Sub sbAplicarCuotaMantenimiento()
Dim strSQL As String, i As Integer
Dim vCodigo As String, vMonto As Currency
Dim vFecha As Date


i = MsgBox("Esta seguro que desea realizar el proceso de la cuota de mantenimiento del Centro de Recreo, se le recuerda" _
                   & " que tiene que ser despues de la aplicacion mensual (Abonos por Planilla)", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


vCodigo = "CMCR"
vMonto = 500
vFecha = fxFechaServidor

lbl.Caption = "Excluyendo Cuotas de Ex-Asociados..."
lbl.Refresh
 
strSQL = "update R set Estado = 'C'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and R.codigo = '" & vCodigo & "' and R.estado = 'A'" _
       & " and S.estadoActual not in('S')"
Call ConectionExecute(strSQL)
 
'ACTUALIZA MONTO CASOS ACTUALES
strSQL = "update R set montoapr = " & vMonto _
       & ",cuota = " & vMonto & ",saldo = " & vMonto _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " where S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and R.codigo = '" & vCodigo & "' and R.estado = 'A'" _
       & " and S.estadoActual = 'S'"
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
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','R'" _
       & ",'N','OT','',0,1,0,'Proceso Automatico Cuota Mantenimiento CR','A'," & fxFechaProcesoSiguiente(GLOBALES.glngFechaCR) _
       & "," & fxFechaProcesoAnterior(GLOBALES.glngFechaCR) & ",'F','AUTOMATICO'" _
       & " from socios where estadoactual = 'S' and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and cedula not in(select cedula from reg_creditos" _
       & " where estado = 'A' and codigo = '" & vCodigo & "'))"
Call ConectionExecute(strSQL)
 

'Cancela Casos en Congelamiento Activado

strSQL = "update R set Estado = 'C'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " where R.estado = 'A' and R.codigo = '" & vCodigo _
       & "' and R.cedula in(select cedula From afi_congelar where estado = 'A'  and fecha_finaliza >= dbo.MyGetdate()" _
       & "  and per_cobro_cuotaCr = 0)"
Call ConectionExecute(strSQL)

'Cancela Casos sin deduccion de Aportes por mas de 2 Meses

strSQL = "update reg_creditos set estado = 'C' where estado = 'A' and codigo = '" & vCodigo _
       & "' and cedula in(select A.cedula From ahorro_consolidado A inner join socios S on A.cedula = S.cedula" _
       & " where S.estadoactual = 'S'  and datediff( m, A.fecAporte, dbo.MyGetdate()) > 2" _
       & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ")"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
lbl.Caption = "Cuota de Mantenimiento Actualizada Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub cmdAplicar_Click()

   Me.Caption = "Cuotas de Mantenimiento Club los Jaúles"
   Call sbAplicarCuotaMantenimiento

End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 10
 

   'Solo Activa las Instituciones CCSS y OPC
    strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX " _
           & " from instituciones where activa = 1 and cod_institucion in(1,2) order by descripcion"

Call sbLlenaCbo(cboInstitucion, strSQL, False, True)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

