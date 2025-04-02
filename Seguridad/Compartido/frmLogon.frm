VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de Sesión ¦ ProGrX: Security System"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9555
   HelpContextID   =   9009
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2340"
   Visible         =   0   'False
   Begin VB.TextBox txt2FA_Metodo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "SMS"
      ToolTipText     =   "Método o Canal para el 2FA"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txt2FA_Codigo 
      Alignment       =   2  'Center
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
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   9
      ToolTipText     =   "Digite su Código de Autenticación"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Digite su contraseña del sistema"
      Top             =   1755
      Width           =   3132
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Digite Aqui su nombre de usuario"
      Top             =   1080
      Width           =   3132
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   480
   End
   Begin VB.TextBox txtBaseDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      ToolTipText     =   "Digite el nombre de la Base de Datos"
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      ToolTipText     =   "Digite el nombre del servidor que tiene la base de datos"
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lbl2FA 
      BackStyle       =   0  'Transparent
      Caption         =   "2FA Doble Factor de Autenticación"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Left            =   5760
      TabIndex        =   8
      Top             =   1515
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4920
      X2              =   9840
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgX 
      Height          =   3300
      Left            =   0
      Picture         =   "frmLogon.frx":030A
      Stretch         =   -1  'True
      Tag             =   "1200"
      Top             =   0
      Width           =   9555
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, vPaso As Boolean
Dim m2FA_Utiliza As Boolean, m2FA_Pass As Boolean




Function fx2FA_GenerarCodigo() As String
    Dim codigo As String
    Randomize
    codigo = Format(Int((999999 - 100000 + 1) * Rnd + 100000), "000000")
    fx2FA_GenerarCodigo = codigo
End Function



Function fx2FA_EnviarCodigoEmail(userEmail As String, codigo As String) As Boolean
    Dim strSQL As String
    
    Dim objEmail As Object
    
    Set objEmail = CreateObject("CDO.Message")
    
    On Error GoTo vError
    
    ' Configuración del servidor SMTP
    With objEmail.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True ' TLS
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "InicioSeguro2FA@ProGrX.info"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "M4t3r@Pl4N#C0m@wK%f1*.!"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = True
        .Update
    End With

    'Registra el Código
    strSQL = "exec sp2FA_Usuario_Codigo '" & glogon.Usuario & "', '" & codigo & "'"
    Call ConectionExecute(strSQL)

    ' Configurar y enviar el mensaje
    With objEmail
        .To = userEmail
        .From = "InicioSeguro2FA@ProGrX.info"
        .Subject = "Código de Verificación"
        .TextBody = "Tu código de verificación es: " & codigo
        .Send
    End With

    fx2FA_EnviarCodigoEmail = True
    
    Exit Function
    
vError:
   fx2FA_EnviarCodigoEmail = False
   'MsgBox Err.Description
    
End Function


Function fx2FA_ValidaCodigo() As Integer
Dim pResult As Integer
Dim strSQL As String, rs As New ADODB.Recordset

pResult = 0

On Error GoTo vError

strSQL = "exec sp2FA_Usuario_Codigo_Valida '" & glogon.Usuario & "', '" & txt2FA_Codigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
    fx2FA_ValidaCodigo = rs!Resultado
rs.Close
Exit Function

vError:
  fx2FA_ValidaCodigo = pResult

End Function

Private Sub sb2FA_Procesa()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pEmail As String, pMovil As String
Dim p2FA_Metodo As String, p2FA_Utiliza As Integer, p2FA_Codigo As String, p2FA_Send As Boolean

On Error GoTo vError

p2FA_Send = False

strSQL = "exec sp2FA_Usuario_Cfg '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    'UserID , TFA_IND, TFA_METODO, EMAIL, TEL_CELL
    pEmail = Trim(rs!EMAIL)
    pMovil = Trim(rs!Tel_Cell)
    p2FA_Utiliza = rs!TFA_IND
    p2FA_Metodo = rs!TFA_METODO
    
    txt2FA_Metodo.Text = p2FA_Metodo
    
    If p2FA_Utiliza = 1 Then
        
        m2FA_Utiliza = True
        
        lbl2FA.Visible = True
        txt2FA_Metodo.Visible = True
        txt2FA_Codigo.Visible = True
        
        txtUserName.Locked = True
        txtPassword.Locked = True
        
        txtUserName.BackColor = RGB(225, 241, 196)
        txtPassword.BackColor = RGB(225, 241, 196)
        
        txt2FA_Codigo.SetFocus
        
        p2FA_Codigo = fx2FA_GenerarCodigo
        txt2FA_Codigo.Tag = p2FA_Codigo
        
        Select Case p2FA_Metodo
            Case "MAIL"
                    p2FA_Send = fx2FA_EnviarCodigoEmail(pEmail, p2FA_Codigo)
            Case "SMS"
            Case "APP"
         End Select
        
    
    End If
End If

'No Utiliza o Existe un Error de aplicación
If p2FA_Utiliza = 0 Then
    Me.MousePointer = vbDefault
    Unload frmLogon
End If


'If p2FA_Utiliza = 1 And p2FA_Send Then
'    Me.MousePointer = vbDefault
'    Unload frmLogon
'End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    End

End Sub


Private Sub sbAceptar()
Dim Resultado   As Integer
Dim strSQL      As String


Me.MousePointer = vbHourglass
    
'Abre Conexión al Portal
Call PGX_OpenPortal_Admin

glogon.Clave = txtPassword.Text
glogon.Usuario = UCase(txtUserName.Text)

glogon.Conectado = 0

Resultado = fxLogonSuccess
    
Me.MousePointer = vbDefault
    
Select Case Resultado
   Case 1 'Error de Ingreso de Aplicacion
   
     MsgBox "No se pudo establecer la conexión con el servidor de la aplicación..."
       
   Case 2 'Error de Credenciales de Usuario

           i = i + 1
    
           If i = glogon.Intentos Then
              MsgBox "No se permiten más intentos...! Su cuenta ha sido bloqueada por " & glogon.TiempoLock & " minutos!", vbExclamation
              
              'Registra el Tiempo de Bloqueo en Log de Transacciones de Cuentas y Registro de Usuarios
'              Call sbSEGCuentaLog("07", "Sobrepasa los (" & glogon.Intentos & ") intentos permitidos, Su cuenta ha sido bloqueda por " & glogon.TiempoLock _
'                                & " minutos!")
              Call sbCancelar
           Else
              MsgBox "El Usuario o Contraseña no fueron encontrados, verifique!"
           End If
       
       Case 0 'Conección Establecida
            glogon.Conectado = 1
'            Call sbSEGCuentaLog("10")
'            Call EscribeArchivoInicio
'
            Me.MousePointer = vbDefault
            Call sb2FA_Procesa
            'Unload frmLogon
    End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCancelar()
  If glogon.Conectado = 0 Then
    End
  Else
'    Call sbSEGCuentaLog("11")
    Unload Me
  End If
End Sub


Private Sub Form_Load()
Dim WS As Object

'Estado de la Versión
'lblVersion.Caption = GLOBALES.SysVersion
m2FA_Utiliza = False
m2FA_Pass = False

lbl2FA.Visible = False
txt2FA_Codigo.Visible = False
txt2FA_Metodo.Visible = False

Set WS = CreateObject("WScript.Network")
glogon.Maquina = WS.ComputerName
glogon.AppName = App.ProductName
glogon.AppVersion = App.Major & "." & App.Minor & "." & Format(App.Revision, "00") & ".r" & GLOBALES.SysVersion

lblVersion.Caption = glogon.AppVersion
 i = 0
 vPaso = False
 
On Error Resume Next
   
  glogon.Conectado = 0
   
  If glogon.Conection.State = 1 Then glogon.Conection.Close
   
   txtUserName.Text = WS.UserName
   
   txtUserName.SelStart = 0
   txtUserName.SelLength = Len(txtUserName)
   
   Call LeeArchivoInicio
    
'
'On Error Resume Next
'   txtUserName.SetFocus

End Sub


Private Sub LeeArchivoInicio()
On Error GoTo error
    Dim vPosIni As Integer
    Dim vPosFin As Integer
    Dim vLinea As String
    Dim vArchivo As String
    Dim vLargoArchivo As Long
    Dim wEmpresa As String
    
    
    vArchivo = App.Path & "\Logon.ini"
    vLargoArchivo = FileLen(vArchivo)
    
    Open vArchivo For Input As #1
        vLinea = Input(vLargoArchivo, #1)
        
        vPosIni = InStr(1, vLinea, "]=", 1)
        vPosIni = vPosIni + 2
        vPosFin = InStr(1, vLinea, ";", 1)
        wEmpresa = Mid(vLinea, vPosIni, vPosFin - vPosIni)
        
        vPosIni = InStr(vPosFin, vLinea, "]=", 1)
        vPosIni = vPosIni + 2
        vPosFin = InStr(vPosIni, vLinea, ";", 1)
        txtServer.Text = Mid(vLinea, vPosIni, vPosFin - vPosIni)
        
        vPosIni = InStr(vPosFin, vLinea, "]=", 1)
        vPosIni = vPosIni + 2
        vPosFin = InStr(vPosIni, vLinea, ";", 1)
        txtBaseDatos.Text = Mid(vLinea, vPosIni, vPosFin - vPosIni)
    Close #1
    
Salir:
    Timer1.Interval = 1
    Exit Sub
    
error:
    If Err.Number = 53 Then
       'Call sbOpciones
    Else
        MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    End If
End Sub

Private Sub EscribeArchivoInicio()
On Error GoTo error
    Dim vArchivo As String
    
    vArchivo = App.Path & "\Logon.ini"
    Open vArchivo For Output As #1
        Write #1, "[Empresa]=NombreDeLaEmpresa;"
        Write #1, "[NombreServidor]=" & txtServer.Text & ";"
        Write #1, "[NombreBaseDatos]=" & txtBaseDatos.Text & ";"
    Close #1
Salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub


Private Sub Form_Unload(Cancel As Integer)


If m2FA_Utiliza And Not m2FA_Pass Then
    glogon.Conectado = 0
End If

If glogon.Conectado = 0 Then Call sbCancelar

End Sub





Private Sub txt2FA_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Select Case fx2FA_ValidaCodigo
        Case 0 'Failt
          MsgBox "Su código es incorrecto, vuelva a intentar!", vbExclamation
        Case 1 'Pass
            m2FA_Pass = True
            Unload Me
        Case 3 'Vencido
          MsgBox "Su código se encuentra vencido, tiene que volver a ingresar!", vbExclamation
          End
    End Select
End If

End Sub

Private Sub txtBaseDatos_GotFocus()
txtBaseDatos.BackColor = txtPassword.BackColor
txtBaseDatos.ForeColor = vbBlue
End Sub

Private Sub txtBaseDatos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call sbAceptar
End Sub


Private Sub txtBaseDatos_LostFocus()
txtBaseDatos.BackColor = &H80000000
txtBaseDatos.ForeColor = vbBlack
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call sbAceptar
End If
End Sub

Private Sub txtServer_GotFocus()
txtServer.BackColor = txtPassword.BackColor
txtServer.ForeColor = vbBlue
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtBaseDatos.SetFocus
End Sub

Private Sub txtServer_LostFocus()
txtServer.BackColor = &H80000000
txtServer.ForeColor = vbBlack
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtPassword.SetFocus
End Sub


