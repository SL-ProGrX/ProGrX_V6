VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCC_ControlPolizas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Polizas (INS)"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5295
   HelpContextID   =   9005
   Icon            =   "CC_ControlPolizas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReportes 
      Caption         =   "Reportes"
      ForeColor       =   &H00C00000&
      Height          =   2415
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdImprime 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdActualizaReportes 
         Caption         =   "Actualiza Reportes"
         Height          =   1335
         Left            =   2280
         Picture         =   "CC_ControlPolizas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "General para INS"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Modificaciones"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Exclusiones"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Inclusiones"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3480
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.OptionButton optPolizas 
      Caption         =   "Reportes"
      Height          =   495
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   2775
   End
   Begin MSComctlLib.ListView lswDetalle 
      Height          =   4815
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6421
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   5640
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CC_ControlPolizas.frx":0614
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4920
      Picture         =   "CC_ControlPolizas.frx":0938
      Stretch         =   -1  'True
      ToolTipText     =   "Actualiza Listado de Codigos para Polizas"
      Top             =   240
      Width           =   255
   End
   Begin VB.Image imgCodigosPSD 
      Height          =   255
      Left            =   4560
      Picture         =   "CC_ControlPolizas.frx":0C42
      Stretch         =   -1  'True
      ToolTipText     =   "Códigos de Polizas Definidos en el sistema"
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      Caption         =   "..."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   4335
   End
End
Attribute VB_Name = "frmCC_ControlPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function fxEsPoliza(strCodigo As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select coalesce(count(*),0) as Existe from catalogo where poliza = 'S' and Codigo = '" & Trim(strCodigo) & "'"
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic

fxEsPoliza = IIf((rsX!existe = 1), True, False)
rsX.Close
End Function

Private Sub sbActualizaINI()
'Dim fn, lng As Long, itmX As ListItem
'
'On Error Resume Next
'
'fn = FreeFile
'
'Kill GLOBALES.gReportes & "\PSD.INI"
'
'Open GLOBALES.gReportes & "\PSD.INI" For Output As #fn   ' Create file name.
'
'
'With lswDetalle
' For lng = 0 To .ListItems.Count
'  If .ListItems.Item(lng).Checked Then
'      Print #fn, .ListItems.Item(lng).Text
'  End If
' Next lng
'End With
'
'Close #fn

End Sub

Private Sub cmdAceptar_Click()

Select Case True
  Case optPolizas(3).Value 'Reportes
   fraReportes.Visible = True
End Select

End Sub


Private Function fxExisteRegistroPoliza(strCedula As String) As Integer
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select coalesce(count(*),0) as Existe,cedula,PSD_actual from control_psd where cedula = '" & strCedula & "' group by cedula,psd_actual"
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic

'0, no existe
'1, existe
'2, reingresa
If Not rsX.EOF And Not rsX.BOF Then
    If rsX!existe = 1 And rsX!psd_actual > 0 Then
     fxExisteRegistroPoliza = 1
    Else
     fxExisteRegistroPoliza = 2
    End If
    If rsX!existe = 0 Then fxExisteRegistroPoliza = 0
Else
    fxExisteRegistroPoliza = 0
End If
rsX.Close
End Function

Private Sub cmdActualizaReportes_Click()
Dim iRespuesta As Integer, strSQL As String, rs As New ADODB.Recordset
Dim lngFecha As Long

'lngFecha = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)

lngFecha = Format(fxFechaServidor, "yyyymm")

iRespuesta = MsgBox("Esta seguro que desea Actualizar la información de los Reportes de Polizas, RECUERDE QUE SE" _
           & " DEBERIA DE REALIZAR SOLO UNA VEZ AL MES", vbYesNo)

If iRespuesta = vbYes Then
 Me.MousePointer = vbHourglass
 
 lblEstado.Caption = "Cargando y Actualizando Info. para Reportes de póliza"
 lblEstado.Refresh
 
 glogon.Conection.Execute "update control_psd set PSD_ANTERIOR = PSD_ACTUAL,PSD_ACTUAL = 0"
 
 strSQL = "SELECT CEDULA,SUM(MONTOAPR) AS MONTO" _
        & " FROM REG_CREDITOS WHERE ESTADO ='A' AND SALDO > 0 AND CODIGO IN(" & fxCodigosPoliza _
        & ") GROUP BY CEDULA"
        
 rs.CursorLocation = adUseServer
 rs.Open strSQL, glogon.Conection, adOpenStatic
 prgBar.Max = rs.RecordCount + 1
 prgBar.Value = 1
 
 Do While Not rs.EOF
  lblEstado.Caption = "Procesando Registro # " & prgBar.Value & " DE " & prgBar.Max
  lblEstado.Refresh
  
  Select Case fxExisteRegistroPoliza(rs!CEDULA)
  Case 0
  'Inserta
    strSQL = "insert control_psd(cedula,psd_anterior,psd_actual,psd_fecha) values('" & Trim(rs!CEDULA) _
           & "',0," & rs!Monto & "," & lngFecha & ")"
  Case 1
   'Actualiza
    strSQL = "update control_psd set psd_Actual = " & rs!Monto _
           & " where cedula = '" & Trim(rs!CEDULA) & "'"
  Case 2
   'Actualiza con reingreso
    strSQL = "update control_psd set psd_Actual = " & rs!Monto _
           & ",psd_fecha = " & lngFecha _
           & " where cedula = '" & Trim(rs!CEDULA) & "'"
  End Select
  
  glogon.Conection.Execute strSQL
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  
  rs.MoveNext
 Loop
 rs.Close
 
 Call Bitacora("Aplica", "Actualización Reportes de Polizas")
 
 Me.MousePointer = vbDefault
 prgBar.Value = 1
 lblEstado.Caption = "..."
 MsgBox "Actualización de Reportes de Pólizas Realizado...", vbInformation
 
End If

End Sub

Private Sub cmdCancelar_Click()
fraReportes.Visible = False
End Sub

Private Sub cmdImprime_Click()
fraReportes.Visible = False
Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
'    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes - Control Pólizas"
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxSIFPathReportes("CrdControlPolizas.rpt")
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "usuario='" & glogon.Usuario & "'"
  
    Select Case True
      Case optReporte(0).Value 'Inclusiones
        .Formulas(3) = "subtitulo='INCLUSIONES'"
        .SelectionFormula = "{CONTROL_PSD.PSD_ANTERIOR}=0 AND {CONTROL_PSD.PSD_ACTUAL}>0"
      Case optReporte(1).Value 'Exclusiones
        .Formulas(3) = "subtitulo='EXCLUSIONES'"
        .SelectionFormula = "{CONTROL_PSD.PSD_ANTERIOR}>0 AND {CONTROL_PSD.PSD_ACTUAL}=0"
      Case optReporte(2).Value 'Modificaciones
        .Formulas(3) = "subtitulo='MODIFICACIONES'"
        .SelectionFormula = "{CONTROL_PSD.PSD_ANTERIOR}<>{CONTROL_PSD.PSD_ACTUAL} AND" _
                          & "{CONTROL_PSD.PSD_ANTERIOR} > 0 AND {CONTROL_PSD.PSD_ACTUAL} > 0"
      Case optReporte(3).Value 'General
        .Formulas(3) = "subtitulo='GENERAL'"
        .SelectionFormula = "{CONTROL_PSD.PSD_ACTUAL}>0"
    End Select
    
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Function fxMarcada(strCodigo As String) As Boolean
'Dim fn, strCadena As String
'
'fn = FreeFile
'On Error GoTo ErrorFx
'fxMarcada = False
'
'Open GLOBALES.gReportes & "\PSD.INI " For Input As #fn   ' Create file name.
'
'Do While Not EOF(fn)
'    Input #fn, strCadena
'    If strCadena = strCodigo Then
'      fxMarcada = True
'      Exit Function
'    End If
'Loop
'Close #fn
'
'ErrorFx:

End Function


Private Function fxCodigosPoliza() As String
'Dim fn, strCadena As String, vPaso As Boolean
'Dim strRes As String
'
'fn = FreeFile
'vPaso = False
'
'On Error GoTo ErrorFx
'
'strRes = ""
'
'Open GLOBALES.gReportes & "\PSD.INI " For Input As #fn   ' Create file name.
'
'Do While Not EOF(fn)
'    Input #fn, strCadena
'
'    If vPaso Then
'       strRes = strRes & ",'" & Trim(strCadena) & "'"
'    Else
'      vPaso = True
'      strRes = "'" & Trim(strCadena) & "'"
'    End If
'Loop
'Close #fn
'
'ErrorFx:
'fxCodigosPoliza = strRes

End Function

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem
Dim fn

 vModulo = 3
 ' Call Formularios(Me)

rs.CursorLocation = adUseServer
strSQL = "select codigo,descripcion,convenio from catalogo where poliza = 'N'"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

With lswDetalle
 .ListItems.Clear
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!Codigo, , 1)
      itmX.Tag = itmX.Index
      itmX.SubItems(1) = rs!Descripcion
      If rs!convenio = "N" Then
        itmX.Checked = True
      Else
        itmX.Checked = fxMarcada(rs!Codigo)
      End If
      rs.MoveNext
 Loop
End With
rs.Close

Call RefrescaTags(Me)

End Sub

Private Sub Image1_Click()
 Call sbActualizaINI
 MsgBox "Nuevo Listado de Códigos para Polizas, registrado...", vbInformation
End Sub

Private Sub imgCodigosPSD_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strCadena As String

strCadena = "CODIGO" & vbTab & "DESCRIPCION" & vbCrLf
strSQL = "select codigo,descripcion from catalogo where poliza = 'S'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 strCadena = strCadena & vbCrLf & rs!Codigo & vbTab & rs!Descripcion
 rs.MoveNext
Loop
rs.Close

MsgBox strCadena, vbInformation, "Códigos de Pólizas Definidos en el Sistema"

End Sub



