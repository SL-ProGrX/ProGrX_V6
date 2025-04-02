VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCR_AcuerdoJDTasas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acuerdo de Junta Directiva s/Tasas de Vivienda"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   Icon            =   "frmCR_AcuerdoJDTasas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_AcuerdoJDTasas.frx":6852
   ScaleHeight     =   6675
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdAcuerdoJDTasas 
      Caption         =   "&Firma Acuerdo de Tasa de Vivienda (JD)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   1320
      Picture         =   "frmCR_AcuerdoJDTasas.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5895
      Width           =   6135
   End
   Begin TabDlg.SSTab ssTabMain 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmCR_AcuerdoJDTasas.frx":138F6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtOperacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDetalle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtDetalle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3585
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   900
         Width           =   6855
      End
      Begin VB.TextBox txtOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   1440
         TabIndex        =   2
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Registro de Acuerdo de Tasas y Plan de Ahorro de la JD "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblAcuerdo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   5640
      Width           =   6135
   End
End
Attribute VB_Name = "frmCR_AcuerdoJDTasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mGarantia As String, mCedula As String, mFecha As Date

Private Sub cmdAcuerdoJDTasas_Click()
Dim strSQL As String

On Error GoTo vError

If lblAcuerdo.Tag = "S" Then
 MsgBox "Ya firmo el Arcuerdo anteriormente...verifique!", vbExclamation
 Exit Sub
End If

If mFecha > CDate("2009/06/18") Then
  MsgBox "No se puede aprobar el acuerdo, porque fue formalizada el...:" & Format(mFecha, "dd/mm/yyyy") & vbCrLf _
        & " [Esta fuera del rango de aprobación del Acuerdo]"
  Exit Sub
End If

strSQL = "update reg_creditos set JD_ACUERDO_TASAS = getdate()" _
       & " where id_solicitud = " & txtOperacion
glogon.Conection.Execute strSQL


strSQL = "insert socios_mensajes(fecha,cedula,usuario,vencimiento,mensaje) values(getdate(),'" _
       & mCedula & "','" & glogon.Usuario & "','2011/03/31','Asociado Firma Ademdum aceptando Trasladar su crédito" _
       & " a Tasa Fija. Según acuerdo de Junta Directiva No.01-1253-09 Fecha de Acuerdo: 18/06/2009')"
glogon.Conection.Execute strSQL

Call Bitacora("Aplica", "Registra el Acuerdo de Tasas de Junta Directiva: " & txtOperacion.Text)

MsgBox "Acuerdo Registrado Satisfactoriamente...", vbInformation

Call txtOperacion_Change

Exit Sub

vError:
 MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Call txtOperacion_Change

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtOperacion_Change()
    txtDetalle.Text = ""
    txtDetalle.Tag = "N"
    lblAcuerdo.Caption = ""
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta

Exit Sub

vError:

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtDetalle.Text = ""
txtDetalle.Tag = "N"


strSQL = "Select R.id_solicitud,R.codigo,R.cedula,S.nombre,R.montoapr,R.plazo,R.int, R.cuota,R.cod_destino" _
       & ",D.descripcion as DestinoX,C.descripcion as LineaX,R.userfor,R.fechaforp,R.observacion,R.garantia,R.JD_ACUERDO_TASAS" _
       & " from reg_creditos R inner join socios S on R.cedula = S.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " left join catalogo_Destinos D on R.cod_destino = D.cod_Destino" _
       & " where R.estado = 'A' and R.id_solicitud = " & txtOperacion.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then

  mGarantia = rs!GARANTIA
  mCedula = Trim(rs!Cedula)
  mFecha = rs!Fechaforp
  
  If IsNull(rs!JD_ACUERDO_TASAS) Then
     lblAcuerdo.Caption = "No a Firmado el Acuerdo de JD para Ajuste de Tasas de Crédito de Vivienda"
     lblAcuerdo.Tag = "N"
  Else
     lblAcuerdo.Caption = "Firmó el Acuerdo el : " & rs!JD_ACUERDO_TASAS
     lblAcuerdo.Tag = "S"
  End If
  
  txtDetalle = txtDetalle & vbCrLf & "# Solicitud : " & vbTab & rs!ID_SOLICITUD
  txtDetalle = txtDetalle & vbCrLf & "Línea       : " & vbTab & rs!Codigo & " - " & rs!LineaX
  txtDetalle = txtDetalle & vbCrLf & "Destino     : " & vbTab & rs!cod_destino & " - " & rs!DestinoX
  txtDetalle = txtDetalle & vbCrLf & "Cédula      : " & vbTab & Trim(rs!Cedula) & " - " & rs!Nombre & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & "Monto       : " & vbTab & Format(rs!montoapr, "Standard")
  txtDetalle = txtDetalle & vbCrLf & "Plazo       : " & vbTab & rs!Plazo
  txtDetalle = txtDetalle & vbCrLf & "Tasa        : " & vbTab & rs!Int
  txtDetalle = txtDetalle & vbCrLf & "Cuota       : " & vbTab & Format(rs!Cuota, "Standard") & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & "Garantía    : " & vbTab & fxGarantia(rs!GARANTIA) & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Fecha       : " & vbTab & Format(rs!Fechaforp, "dd/mm/yyyy")
  txtDetalle = txtDetalle & vbCrLf & "Usuario     : " & vbTab & rs!Userfor & vbCrLf

  txtDetalle = txtDetalle & vbCrLf & "Notas : " & rs!observacion & ""

  txtDetalle.Tag = "S"
End If
rs.Close

Me.MousePointer = vbDefault

If txtDetalle.Tag = "N" Then
   txtDetalle.ForeColor = vbRed
   MsgBox " La Solicitud no cumple con alguno(s) de los siguientes parámetros:" _
          & vbCrLf & " 1. No se encuentra Formalizada" & vbCrLf & " 2. No Existe la Solicitud (Credito)" _
          , vbExclamation
Else
   txtDetalle.ForeColor = vbBlue
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 Call txtOperacion_Change

End Sub



