VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmLogon_Datos_Update 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizar los datos para Notificaciones del Usuario:"
   ClientHeight    =   3660
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5640
      TabIndex        =   2
      Top             =   1875
      Width           =   3855
   End
   Begin VB.TextBox txtMovil 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   120
      Top             =   600
   End
   Begin XtremeSuiteControls.PushButton btnActualiza 
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
      _Version        =   1245185
      _ExtentX        =   4048
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Actualiza datos de Contacto"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmLogon_Datos_Update.frx":0000
      ImageAlignment  =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   1635
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Móvil"
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
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image imgX 
      Height          =   3300
      Left            =   0
      Picture         =   "frmLogon_Datos_Update.frx":07DE
      Stretch         =   -1  'True
      Tag             =   "1200"
      Top             =   0
      Width           =   9555
   End
End
Attribute VB_Name = "frmLogon_Datos_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mUserId As Long

Private Sub sbLeeDatosUsuario()

On Error GoTo vError

strSQL = "exec  spSEG_Logon_Info '" & glogon.Usuario & "','" & glogon.Maquina_MAC & "'"
 
Call OpenRecordSet(rs, strSQL, 1)

txtMovil.Text = Trim(rs!Tel_Cell)
txtEmail.Text = Trim(rs!Email)

'El UserID viene alterado por 7 como medida de seguridad
mUserId = rs!UserId / 7
 
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnActualiza_Click()
Dim vEmail_Valida As Boolean

txtMovil.Text = Replace(txtMovil.Text, "-", "")
txtMovil.Text = Replace(txtMovil.Text, " ", "")

If Not IsNumeric(txtMovil.Text) Then
  MsgBox "Número de Teléfono Móvil no es válido, verifique!", vbExclamation
  Exit Sub
End If

vEmail_Valida = False

If InStr(1, txtEmail.Text, "@") > 0 Then
    vEmail_Valida = True
End If

If InStr(1, txtEmail.Text, ".") = 0 And vEmail_Valida = True Then
    vEmail_Valida = False
End If

If Not vEmail_Valida Then
   MsgBox "Correo Electrónico no es válido, verifique!", vbExclamation
   Exit Sub
End If


strSQL = "exec spSEG_Logon_Info_Update '" & glogon.Usuario & "','" & txtEmail.Text & "','" & txtMovil.Text & "'," & mUserId

Call ConectionExecute(strSQL, 1)

If Not glogon.error Then
    MsgBox "Información de Contacto del Usuario actualizada satisfactoriamente!", vbInformation
    Unload Me
End If

End Sub

Private Sub Form_Load()

Me.BackColor = RGB(70, 111, 178)
Me.ScaleMode = 3

Call sbLeeDatosUsuario

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
'
'imgX.Width = Me.ScaleWidth
'imgX.Height = Me.ScaleHeight

End Sub
