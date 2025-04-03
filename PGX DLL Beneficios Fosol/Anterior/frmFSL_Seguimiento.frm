VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Seguimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de Expedientes"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Seguimiento"
      TabPicture(0)   =   "frmFSL_Seguimiento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tlbGuarda"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboGestion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtComunicado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Historico"
      TabPicture(1)   =   "frmFSL_Seguimiento.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vGridHistoricoGestiones"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtComunicado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   7215
      End
      Begin VB.ComboBox cboGestion 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   4815
      End
      Begin MSComctlLib.Toolbar tlbGuarda 
         Height          =   360
         Left            =   8160
         TabIndex        =   12
         Top             =   4740
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               ImageIndex      =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread vGridHistoricoGestiones 
         Height          =   4620
         Left            =   -74760
         TabIndex        =   13
         Top             =   420
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   8149
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   4
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Seguimiento.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Comunicado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Gestión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Exposición a Riesgo de la persona"
         Top             =   600
         Width           =   5235
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   5445
      End
      Begin VB.TextBox txtCedula 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   240
         Width           =   2070
      End
      Begin VB.TextBox txtExpediente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Expediente"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Asociado"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Expediente"
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
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   11040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":06AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":6F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":D771
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":13FD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":1A835
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":21097
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Seguimiento.frx":36209
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFSL_Seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vConsecutivo As Long, vExpediente As Long, vGestion As Integer, vObservaciones As String
Dim vFecha As String, vUsuario As String

Private Sub Form_Activate()
  vModulo = 1
End Sub

Private Sub Form_Load()
  vModulo = 1
  ssTab.Tab = 0
  If GLOBALES.gTag <> Empty And GLOBALES.gTag2 <> Empty And GLOBALES.gTag3 <> Empty Then
    txtCedula.Text = GLOBALES.gTag
    txtNombre.Text = GLOBALES.gTag2
    txtExpediente.Text = GLOBALES.gTag3
  End If
  Call sbCargaGestiones
  GLOBALES.gTag = ""
  GLOBALES.gTag2 = ""
  GLOBALES.gTag3 = ""
  vFecha = Format(fxFechaServidor, "yyyymmdd")
  vUsuario = glogon.Usuario
End Sub

Private Sub sbCargaGestiones()
On Error GoTo vError

cboGestion.Clear

strSQL = "Select COD_GESTION, DESCRIPCION, ESTADO from FSL_GESTIONES"
rs.Open strSQL, glogon.Conection, adOpenStatic
  
  Do While Not rs.EOF
     cboGestion.AddItem (rs!COD_GESTION & " - " & Trim(rs!Descripcion))
     rs.MoveNext
  Loop
    
  If rs.RecordCount > 0 Then
     rs.MoveFirst
     cboGestion.Text = rs!COD_GESTION & " - " & rs!Descripcion
  End If

  rs.Close

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub sbGuardarSeguimiento()
On Error GoTo vError
 
strSQL = "Insert FSL_EXPEDIENTE_GESTIONES (CONSECUTIVO, COD_EXPEDIENTE, COD_GESTION, OBSERVACIONES" _
       & ", USUARIO_REGISTRA, FECHA_REGISTRO) values (" & vConsecutivo & "," & vExpediente & "," & vGestion & ",'" & vObservaciones & "'" _
       & ", '" & vUsuario & "','" & vFecha & "' )"
glogon.Conection.Execute strSQL

MsgBox "La Gestión se guardo correctamente.", vbInformation

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub sbTraerHistorico()
Dim i As Integer
On Error GoTo vError

strSQL = "Select EG.CONSECUTIVO, EG.COD_EXPEDIENTE, EG.COD_GESTION,G.DESCRIPCION, EG.OBSERVACIONES " _
       & ", EG.USUARIO_REGISTRA , EG.FECHA_REGISTRO " _
       & " from FSL_EXPEDIENTE_GESTIONES EG " _
       & " inner join FSL_GESTIONES G on EG.COD_GESTION = G.COD_GESTION " _
       & " Where EG.COD_EXPEDIENTE = " & txtExpediente.Text & ""
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

'Ojo revisar para cargar el detalle de la gestion
With vGridHistoricoGestiones
 .MaxRows = 1
 .Row = .MaxRows
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
 Next i

 Do While Not rs.EOF
  .Row = .MaxRows
  .Col = 1
  .Text = rs!FECHA_REGISTRO
      
  .Col = 2
  .Text = rs!Descripcion
  
  .Col = 3
  .AutoSize = True
  .Text = rs!observaciones
  
  .Col = 4
  .Text = rs!USUARIO_REGISTRA
      
  rs.MoveNext
  .MaxRows = .MaxRows + 1
 Loop
End With
rs.Close
Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
  If ssTab.Tab = 1 Then
   If Not IsNumeric(txtCedula.Text) Or txtNombre.Text = Empty Then Exit Sub
   Call sbTraerHistorico
  End If
End Sub

Private Sub tlbGuarda_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
 
 If txtComunicado.Text = Empty Then
   MsgBox "Se debe de incluir un detalle para almacenar la gestión", vbCritical
   Exit Sub
 End If
 
 strSQL = "Select coalesce(max(CONSECUTIVO),0) + 1 as 'Consecutivo' from FSL_EXPEDIENTE_GESTIONES"
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
  vConsecutivo = rs!consecutivo
 rs.Close
 
 vExpediente = txtExpediente.Text
 vGestion = SIFGlobal.fxSIFCodText(cboGestion)
 vObservaciones = txtComunicado.Text
 
 
 Call sbGuardarSeguimiento
 Call Bitacora("Registra", "Registra gestion realizada a Expediente Fosol #" & txtExpediente.Text)
 
Exit Sub
vError:
   MsgBox Err.Description, vbExclamation
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion
vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select coalesce(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then
   rs.Close
   strSQL = "select cedula from reg_creditos where id_solicitud = " & txtCedula
   rs.Open strSQL, glogon.Conection, adOpenStatic
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close

If vCedTemp = "" Then
  Call sbConsulta(txtCedula.Text)
Else
  Call sbConsulta(vCedTemp)
End If
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda
txtExpediente.Text = Empty
txtComunicado.Text = Empty
Exit Sub
vError:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
Dim vEstadoSocio As String
     
vFianzas = False
    
strSQL = "select S.cedula as CedulaX,S.nombre" _
       & " from socios S " _
       & " Where S.cedula = '" & Trim(pCedula) & "'"

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
 
If Not rs.EOF And Not rs.BOF Then
   
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
   rs.Close
       
   ssTab.Tab = 0
 
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
  
End Sub

Private Sub sbBusqueda()
On Error GoTo vError
  gBusquedas.Convertir = "N"
  gBusquedas.Consulta = "Select cedula,nombre from SOCIOS"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  frmBusquedas.Show vbModal
  txtCedula = Trim(gBusquedas.Resultado)
  gBusquedas.Consulta = ""
  gBusquedas.Columna = ""
  gBusquedas.Orden = ""
  gBusquedas.Resultado = ""
  If Trim(txtCedula) <> "" Then
       Call sbConsulta(txtCedula)
  End If
Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub
