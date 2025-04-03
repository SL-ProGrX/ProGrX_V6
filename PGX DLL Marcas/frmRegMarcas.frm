VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmRegMarcas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Marcas"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   2400
   End
   Begin XtremeSuiteControls.PushButton cmdRegistrar 
      Height          =   852
      Left            =   2760
      TabIndex        =   8
      Top             =   3120
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Registro de Marca "
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRegMarcas.frx":0000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Marca"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   1800
      TabIndex        =   9
      Top             =   360
      Width           =   3612
   End
   Begin VB.Image imgBanner 
      Height          =   1116
      Left            =   0
      Top             =   0
      Width           =   10920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   2172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   2172
   End
   Begin VB.Label lblNombreUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label lblRegUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   2172
   End
End
Attribute VB_Name = "frmRegMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vMarca As Integer
Dim vCodigoHorario As String
Dim WS As Object, vPC As String

Private Sub cmdRegistrar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
strSQL = "insert into MARCAS_REGISTRO(cod_horario,usuario,estacion,fecha,tipo_marca)" _
        & "values('" & vCodigoHorario & "','" & glogon.Usuario & "','" & vPC & "',dbo.MyGetdate()," _
        & " " & vMarca & ")"
Call ConectionExecute(strSQL)

MsgBox "Marca registrada satisfactoriamente"
UnLoad Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 21
End Sub

Private Sub Form_Load()
vModulo = 21


Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

Set WS = CreateObject("WScript.Network")

lblRegUser.Caption = glogon.Usuario
lblFecha.Caption = Format(fxFechaServidor(), "dd/mm/yyyy")
lblHora.Caption = Format(fxFechaServidor(), "hh:mm:ss AMPM")
lblNombreUser = fxNombreUsuario(glogon.Usuario)

vPC = WS.computername
vMarca = fxTipoMarca

Select Case vMarca
  Case 0
    MsgBox "Este usuario no tiene asignado horario"
    UnLoad Me
  Case 5
    MsgBox "Ya cumplio con el las marcas del día"
    UnLoad Me
  Case Else
    Call Formularios(Me)
    Call RefrescaTags(Me)
End Select
  

End Sub



Private Function fxNombreUsuario(vUsuario As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select descripcion from usuarios where nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Or Not rs.BOF Then
  fxNombreUsuario = rs!Descripcion
Else
  fxNombreUsuario = Empty
End If

rs.Close

End Function


Private Function fxTipoMarca() As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoMarca As Integer



'Consulta el codigo del horario que corresponde para la marca
strSQL = "select cod_horario from MARCAS_HORARIOS_USERS where usuario = '" & glogon.Usuario & "' "
Call OpenRecordSet(rs, strSQL)


If Not rs.EOF Or Not rs.BOF Then
   vCodigoHorario = rs!COD_HORARIO
Else
   fxTipoMarca = 0
   Exit Function
End If
rs.Close


'Verifica la última marca realizada
strSQL = "select max(tipo_marca) as marca from marcas_registro where usuario  = '" & glogon.Usuario & "' " _
       & " and fecha between '" & Format(fxFechaServidor, "yyyymmdd 00:00:00.00") & "' and '" & Format(fxFechaServidor, "yyyymmdd 23:59:59.00") & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Or Not rs.BOF Then
  If IsNull(rs!marca) Then 'si la marca es nula, se considera la primera del día
    vTipoMarca = 0
  Else
    vTipoMarca = rs!marca
  End If
End If

rs.Close

strSQL = "select salalmuerzo,entalmuerzo from marcas_horarios where cod_horario = '" & vCodigoHorario & "'"
Call OpenRecordSet(rs, strSQL)


Select Case vTipoMarca
  Case 0
    fxTipoMarca = 1
  Case 1
    If rs!salalmuerzo = 1 Then
      fxTipoMarca = 2
    Else
      fxTipoMarca = 4
    End If
  Case 2
    If rs!entalmuerzo = 1 Then
      fxTipoMarca = 3
    Else
      fxTipoMarca = 4
    End If
 Case 3
     fxTipoMarca = 4
 Case 4
    fxTipoMarca = 5
End Select

End Function


Private Sub Timer1_Timer()
lblHora.Caption = Time
End Sub
