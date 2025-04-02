VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCC_Polizas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Polizas"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "CC_Polizas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      Caption         =   "Configuración de Polizas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   735
      End
      Begin VB.Image imgCerrar 
         Height          =   255
         Left            =   5640
         Picture         =   "CC_Polizas.frx":6852
         Stretch         =   -1  'True
         ToolTipText     =   "Cerrar  y Guardar Configuracion"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "En Base al"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Monto por Millon"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Línea"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   360
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   1755
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "CC_Polizas.frx":6979
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgConfiguracion 
      Height          =   255
      Left            =   7320
      Picture         =   "CC_Polizas.frx":6BE1
      Stretch         =   -1  'True
      ToolTipText     =   "Ver Configuración"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl 
      Caption         =   "Actualización de Polizas de Cartera, Aplicar antes del envio de la planilla."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
End
Attribute VB_Name = "frmCC_Polizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxFondoSolidario(curMonto As Currency, Optional curMontoBase As Currency = 150) As Currency
fxFondoSolidario = (curMonto / 1000000) * curMontoBase
End Function


Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCodigo As String


i = MsgBox("Esta seguro que desea realizar el calculo de las polizas, se le recuerda" _
                   & " que tiene que ser despues de la aplicacion mensual (Abonos por Planilla)", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información..."
lbl.Refresh

vCodigo = txtCodigo.Text

'Eliminar o Cancelar la Deducciones a los casos que ya cancelaron sus prestamos
strSQL = "select X.id_solicitud,X.codigo,X.cedula from reg_creditos X inner join Socios Y on X.cedula = Y.cedula" _
       & " where X.estado = 'A' and X.codigo = '" & vCodigo & "' and Y.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and X.cedula not in(select R.cedula" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where C.retencion = 'N' and C.poliza = 'N'" _
       & " and R.saldo > 0 and R.estado = 'A' and R.proceso <> 'J'" _
       & " group by R.cedula)"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
 lbl.Caption = "Cancelando POLIZAS sin Garantias..." & vbCrLf & vbCrLf & "Procesando " & prgBar.Value & " de " & prgBar.Max
 lbl.Refresh
 
 strSQL = "update reg_creditos set estado = 'C' where id_solicitud = " & rs!Id_solicitud
 glogon.Conection.Execute strSQL
 
 prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close

Call sbFondoSolidarioPaso1

prgBar.Value = 1

Me.MousePointer = vbDefault
lbl.Caption = "Polizas Actualizadas Satisfactoriamente..."

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbExclamation

End Sub


Private Sub sbFondoSolidarioPaso1()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset, iRespuesta As Integer
Dim vCodigo As String, vFecha As Date
Dim vFND As Currency, vFechaProX As Long
Dim vGarantia As String, vComite As Integer


Me.MousePointer = vbHourglass
lbl.Alignment = vbLeftJustify


lbl.Caption = "Cargando Información..."
lbl.Refresh

vCodigo = txtCodigo.Text

vGarantia = fxCrdGarantiaOmision(vCodigo)
vComite = fxCrdIdComiteLinea(vCodigo)


vFecha = fxFechaServidor
vFechaProX = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

'Actualizando FND
strSQL = "select R.cedula,sum(montoapr) as Monto" _
       & " from reg_Creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula " _
       & " where C.retencion = 'N' and C.poliza = 'N' and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and R.saldo > 0 and R.estado = 'A' and R.proceso <> 'J'" _
       & " group by R.cedula"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1


Do While Not rs.EOF
 
 lbl.Caption = "Actualizando ..." & vbCrLf & vbCrLf & "Procesando " & prgBar.Value & " de " & prgBar.Max
 lbl.Refresh
 
 If rs!Monto > 12000000 Then
     vFND = fxFondoSolidario(12000000, CCur(txtMonto))
  Else
     vFND = fxFondoSolidario(rs!Monto, CCur(txtMonto))
 End If
 
 strSQL = "Select id_solicitud,cuota from reg_creditos where codigo = '" & vCodigo & "' and cedula = '" _
        & rs!Cedula & "' and estado = 'A'"
 rsTmp.Open strSQL, glogon.Conection, adOpenStatic
 If rsTmp.EOF And rsTmp.BOF Then
   'Insertar nueva cuota
     strSQL = "insert reg_creditos(id_comite,codigo,cedula,montoapr,plazo,int,interesv" _
          & ",saldo,interesc,amortiza,cuota,prideduc,fecult,estadosol,estado,fechasol,fechaforp" _
          & ",observacion,garantia,tdocumento,ndocumento,tesoreria,userrec,userfor)" _
          & "values(" & vComite & ",'" & vCodigo & "','" & Trim(rs!Cedula) & "'," & vFND & ",999,0,0," & vFND _
          & ",0,0," & vFND & "," & vFechaProX & "," & GLOBALES.glngFechaCR & ",'F','A','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Format(vFecha, "yyyy/mm/dd") & "'" _
          & ",'GENERACION AUTO. POLIZA: " & vFecha & "','" & vGarantia & "','OT','','" & Format(vFecha, "yyyy/mm/dd") _
          & "','" & Trim(glogon.Usuario) & "','" & Trim(glogon.Usuario) & "')"
      glogon.Conection.Execute strSQL

 Else
   If Abs(vFND - rsTmp!Cuota) > 1 Then
     'Actualizar la cuota
     strSQL = "update reg_creditos set cuota = " & vFND & ", saldo = " & vFND & ",montoapr = " & vFND _
            & " where id_solicitud = " & rsTmp!Id_solicitud
     glogon.Conection.Execute strSQL
   End If
 End If
 rsTmp.Close
 
 prgBar.Value = prgBar.Value + 1
 
 rs.MoveNext
Loop
rs.Close


prgBar.Value = 1

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbExclamation

End Sub


Private Sub Form_Load()
Dim strSQL As String

 vModulo = 3
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 cbo.Clear
 cbo.AddItem "01 - Sobre Monto Aprobado"
 cbo.AddItem "02 - Sobre Saldos"
 cbo.Text = "01 - Sobre Monto Aprobado"
 
 txtCodigo.Text = "POL1"
 
 txtMonto.Text = 430

strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones where activa = 1 order by descripcion"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)
 

End Sub

Private Sub imgCerrar_Click()
fra.Visible = False
End Sub

Private Sub imgConfiguracion_Click()
fra.Visible = True
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  gBusquedas.Columna = "Codigo"
  gBusquedas.Consulta = "select Codigo,Descripcion from catalogo"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = " and retencion = 'S' or poliza = 'S'"
  gBusquedas.Orden = "codigo"
  
  frmBusquedas.Show vbModal
  
  txtCodigo = gBusquedas.Resultado
  txtDescripcion = gBusquedas.Resultado2
  
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Consulta = "select Codigo,Descripcion from catalogo"
  gBusquedas.Convertir = "N"
  gBusquedas.Filtro = " and retencion = 'S' or poliza = 'S'"
  gBusquedas.Orden = "descripcion"
  
  frmBusquedas.Show vbModal
  
  txtCodigo = gBusquedas.Resultado
  txtDescripcion = gBusquedas.Resultado2
  
End If

End Sub
