VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCC_EstadoCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados de Cuenta"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   HelpContextID   =   9007
   Icon            =   "CC_EstadoCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Picture         =   "CC_EstadoCuenta.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Estado de Cuenta"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtPeriodoHasta 
      Height          =   315
      Left            =   4560
      TabIndex        =   16
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtPeriodoDe 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdReporteExcedentes 
      Caption         =   "Excedentes"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Picture         =   "CC_EstadoCuenta.frx":685E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Reporte del Estado de Cuenta"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cboSeccion 
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1680
      Width           =   4455
   End
   Begin VB.ComboBox cboDept 
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1320
      Width           =   4455
   End
   Begin VB.ComboBox cbo 
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   4455
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   6
      Top             =   3270
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ComboBox cboSegmento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox chkSegmentos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "General Segmentado por"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox chkSalida 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Salida a Impresora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txtCedula 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Nombre Completo del Socio (Apellidos y Nombre)"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   6600
      X2              =   6600
      Y1              =   2040
      Y2              =   3480
   End
   Begin VB.Label lblZ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sección"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Departamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblEstado 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image imgConstancia 
      Height          =   255
      Left            =   2880
      Picture         =   "CC_EstadoCuenta.frx":D0B0
      Stretch         =   -1  'True
      ToolTipText     =   "Imprime Constancia "
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCC_EstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Or cbo.ListCount <= 0 Then Exit Sub

vPaso = True
    strSQL = "select cod_departamento + ' - ' + descripcion as ItmX from AFdepartamentos where cod_institucion = " & cbo.ItemData(cbo.ListIndex)
    Call sbLlenaCbo(cboDept, strSQL, True, False)
vPaso = False

End Sub



Private Sub cboDept_Click()
Dim strSQL As String

If vPaso Or cboDept.ListCount <= 0 Then Exit Sub

vPaso = True
    strSQL = "select cod_seccion + ' - ' + descripcion as ItmX from AFSecciones where cod_institucion = " & cbo.ItemData(cbo.ListIndex) _
           & " and cod_departamento = '" & fxCodigoCbo(cboDept) & "'"
    Call sbLlenaCbo(cboSeccion, strSQL, True, False)
vPaso = False

End Sub

Private Sub chkSegmentos_Click()

If chkSegmentos.Value = vbUnchecked Then
  txtCedula.Enabled = True
  txtNombre.Enabled = True
  chkSegmentos.Caption = "Individual"
  cboSegmento.Visible = False
  
  
  Me.Height = 2670
  Line2.Y1 = 840
  Line2.Y2 = 840
  Line4.Y1 = 840
  
  cmdReporte.Top = 960
  cmdReporteExcedentes.Top = cmdReporte.Top
  
  Label4.Top = 960
  txtPeriodoDe.Top = 1200
  txtPeriodoHasta.Top = txtPeriodoDe.Top
  
  lblEstado.Top = 960
  
  
Else
  txtCedula.Enabled = False
  txtNombre.Enabled = False
  chkSegmentos.Caption = "General Segmentado por"
  cboSegmento.Visible = True
 
  Me.Height = 3855
  Line2.Y1 = 2040
  Line2.Y2 = 2040
  Line4.Y1 = 2040
  
  cmdReporte.Top = 2160
  cmdReporteExcedentes.Top = cmdReporte.Top
  
  Label4.Top = 2160
  txtPeriodoDe.Top = 2400
  txtPeriodoHasta.Top = txtPeriodoDe.Top
  
  lblEstado.Top = 2280
  
  
End If

cbo.Visible = cboSegmento.Visible
cboSeccion.Visible = cboSegmento.Visible
cboDept.Visible = cboSegmento.Visible

lblX.Visible = cboSegmento.Visible
lblY.Visible = lblX.Visible
lblZ.Visible = lblX.Visible



End Sub


Private Sub cmdReporte_Click()
Dim strSQL As String

If chkSegmentos.Value = vbUnchecked Then
  Call sbEstadoCuenta(txtCedula, chkSalida.Value)
Else
  
  strSQL = ""
  Select Case Mid(cboSegmento.Text, 1, 2)
     Case "01" 'Socios Activos
        strSQL = "{vSIF_EC_Principal.ESTADOACTUAL} = 'S'"
     Case "02" 'Ex-Socios Internos
        strSQL = "{vSIF_EC_Principal.ESTADOACTUAL} = 'A'"
     Case "03" 'Ex-Socios Patronal
        strSQL = "{vSIF_EC_Principal.ESTADOACTUAL} = 'P'"
     Case "04" 'No Socios
        strSQL = "{vSIF_EC_Principal.ESTADOACTUAL} = 'N'"
     Case "05" 'Todos
  End Select
  
  If cbo.Text <> "TODAS" Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{vSIF_EC_Principal.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex)
  End If
  
  If Not (cboDept.Text = "TODOS" Or cboDept.Text = "") Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{vSIF_EC_Principal.COD_DEPARTAMENTO} = '" & fxCodigoCbo(cboDept) & "'"
  End If
  
  If Not (cboSeccion.Text = "TODOS" Or cboSeccion.Text = "") Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{vSIF_EC_Principal.COD_SECCION} = '" & fxCodigoCbo(cboSeccion) & "'"
  End If
  
  
  
  Call sbEstadoCuentaInst(chkSalida.Value, strSQL)
 
End If

End Sub

Private Sub sbBusqueda(Index As Integer)

gBusquedas.Convertir = "N"

Select Case Index
   
   Case 0
       gBusquedas.Resultado = Trim(txtCedula)
       gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
       gBusquedas.Columna = "Cedula"
       gBusquedas.Orden = "Cedula"
       frmBusquedas.Show vbModal
       GLOBALES.gblnBuscando = True
       txtCedula = Trim(gBusquedas.Resultado)
   
   Case 2
       gBusquedas.Resultado = Trim(txtCedula)
       gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
       gBusquedas.Columna = "Nombre"
       gBusquedas.Orden = "Nombre"
       frmBusquedas.Show vbModal
       GLOBALES.gblnBuscando = True
       txtCedula = Trim(gBusquedas.Resultado)

End Select

End Sub


Private Sub cmdReporteExcedentes_Click()
Dim strRuta As String

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Cuentas Corrientes"
     
     .Connect = glogon.ConectRPT
     
    If chkSegmentos.Value = vbChecked Then
     .ReportFileName = fxSIFPathReportes("SIFEstadoExcedentesUnidad.rpt")
     .Formulas(0) = "CORTE='" & txtPeriodoDe & "-" & txtPeriodoHasta & "'"
     .Formulas(1) = "SistemaFecha = 'Fecha/Hora : " & fxFechaServidor & "'"
     .Formulas(2) = "SistemaUsuario = 'Usuario : " & glogon.Usuario & "'"
     
     .SelectionFormula = "{SOCIOS.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex) & " AND " _
        & "{EXC_CIERRE.PERIODO_DE} = " & txtPeriodoDe & " AND " _
        & "{EXC_CIERRE.PERIODO_HASTA} = " & txtPeriodoHasta
     
     .SubreportToChange = "Exc_Carga"
     .SelectionFormula = "{EXC_CARGA.PERIODO_DE} = " & txtPeriodoDe & " AND " _
        & "{EXC_CARGA.PERIODO_HASTA} = " & txtPeriodoHasta & " AND " _
        & "{EXC_CARGA.CEDULA} = {?Pm-EXC_CIERRE.CEDULA}"
    
    
    Else
     
     .ReportFileName = fxSIFPathReportes("SIFEstadoExcedentes.rpt")
     
     .Formulas(0) = "CORTE='" & txtPeriodoDe & "-" & txtPeriodoHasta & "'"
     .Formulas(1) = "SistemaFecha = 'Fecha/Hora : " & fxFechaServidor & "'"
     .Formulas(2) = "SistemaUsuario = 'Usuario : " & glogon.Usuario & "'"
     
     .SelectionFormula = "{EXC_CIERRE.PERIODO_DE} = " & txtPeriodoDe & " AND " _
        & "{EXC_CIERRE.PERIODO_HASTA} = " & txtPeriodoHasta & " AND " _
        & "{EXC_CIERRE.CEDULA} ='" & txtCedula & "'"
     
     .SubreportToChange = "Exc_Carga"
     .SelectionFormula = "{EXC_CARGA.PERIODO_DE} = " & txtPeriodoDe & " AND " _
        & "{EXC_CARGA.PERIODO_HASTA} = " & txtPeriodoHasta & " AND " _
        & "{EXC_CARGA.CEDULA} = {?Pm-EXC_CIERRE.CEDULA}"
    
     Call Bitacora("Imprime", "Estado Excedentes Ced." & txtCedula & " Per." & txtPeriodoDe & "-" & txtPeriodoHasta)
    
    End If
    
     If chkSalida.Value = vbChecked Then .Destination = crptToPrinter
     .PrintReport


End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMes As Integer, vFecha As Date

On Error GoTo vError

Me.Icon = MDIPrincipal.Icon

vModulo = 10


vPaso = True
strSQL = "select cod_institucion as Idx,descripcion as ItmX from instituciones where Activa = 1"
Call sbLlenaCbo(cbo, strSQL, True, True)
vPaso = False


With cboSegmento
  .AddItem "01 - Socios Activos"
  .AddItem "02 - Ex-Socio Interno"
  .AddItem "03 - Ex-Socio Patronal"
  .AddItem "04 - No Socio"
  .AddItem "05 - Todos"
  .Text = "01 - Socios Activos"
End With

Call chkSegmentos_Click


strSQL = "select coalesce(max(periodo_de),0) as PeriodoI from excedentes_parcierre"
rs.Open strSQL, glogon.Conection, adOpenStatic

txtPeriodoDe = rs!periodoi
rs.Close


If txtPeriodoDe = 0 Then
    vFecha = fxFechaServidor
    vMes = Month(vFecha)
    
    If vMes > 9 Then
      txtPeriodoDe = Year(vFecha)
    Else
      txtPeriodoDe = Year(vFecha) + 1
    End If
End If


vError:


End Sub


Private Sub imgConstancia_Click()

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Cuentas Corrientes"
     
     .Connect = glogon.ConectRPT
     
     .Formulas(0) = "formula4 = 'La Sección de Servicio al Asociado de la " & Trim(GLOBALES.gstrNombreEmpresa) _
                              & ", hace constar lo siguiente:'"
     
     .ReportFileName = fxSIFPathReportes("SIFEstadoConstancia.rpt")
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & txtCedula & "'"
     
     .SubreportToChange = "sbCreditos"
     .ReplaceSelectionFormula "{REG_CREDITOS.CEDULA} = {?Pm-SOCIOS.CEDULA} and ({REG_CREDITOS.SALDO} > 0 and {REG_CREDITOS.ESTADO} = 'A' AND {CATALOGO.RETENCION} = 'N')"

     
'        & " and{REG_CREDITOS.ESTADO} = 'A'"
     If chkSalida.Value = vbChecked Then .Destination = crptToPrinter
     .PrintReport
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtCedula_Change()
 txtNombre = fxNombre(txtCedula)
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
KeyAscii = (Validacion(KeyAscii))

If KeyAscii = vbKeyReturn Then
   cmdReporte.SetFocus
End If
End Sub


Private Sub txtCedula_LostFocus()
Dim rs As New ADODB.Recordset

On Error Resume Next

If Trim(txtCedula.Text) <> "" Then
 rs.Source = "Select * from Socios Where Cedula ='" & Trim(txtCedula.Text) & "'"
 rs.Open , glogon.Conection, adOpenStatic
 
 If rs.EOF And rs.BOF Then
    MsgBox "No se encontró registro", vbExclamation
    txtCedula = ""
    txtNombre = ""
    txtCedula.SetFocus
 Else
    txtNombre = ""
    txtNombre = Trim(rs!Nombre)
 End If
 
 rs.Close

Else
   txtNombre = ""
End If

End Sub


Private Sub txtNombre_Change()
txtNombre = fxNombre(txtCedula)
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cmdReporte.SetFocus
End If
End Sub

Private Sub txtPeriodoDe_Change()
On Error Resume Next
 txtPeriodoHasta.Text = Val(txtPeriodoDe.Text) + 1
End Sub

Private Sub txtPeriodoHasta_Change()
On Error Resume Next
 txtPeriodoDe.Text = Val(txtPeriodoHasta.Text) - 1

End Sub
