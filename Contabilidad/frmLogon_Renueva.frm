VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmLogon_Renueva 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso: Solicita Renovación de Cuenta"
   ClientHeight    =   3324
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9828
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   9828
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox_Activa 
      Height          =   1932
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   7212
      _Version        =   1245185
      _ExtentX        =   12721
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Ingrese su Usuario y Token recibido en su correo:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Token_Usuario 
         Height          =   312
         Left            =   4680
         TabIndex        =   10
         Top             =   480
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Token 
         Height          =   312
         Left            =   4680
         TabIndex        =   11
         Top             =   840
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnToken_Activa 
         Height          =   492
         Left            =   4680
         TabIndex        =   12
         Top             =   1320
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Renovar Contraseña"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmLogon_Renueva.frx":0000
         ImageAlignment  =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   3480
         TabIndex        =   14
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Token:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   3480
         TabIndex        =   13
         Top             =   840
         Width           =   2052
      End
   End
   Begin XtremeSuiteControls.RadioButton RbtToken 
      Height          =   852
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2052
      _Version        =   1245185
      _ExtentX        =   3619
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Solicitar el TOKEN para Renovar mi contraseña"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      Appearance      =   2
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox_Solicita 
      Height          =   1932
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   7212
      _Version        =   1245185
      _ExtentX        =   12721
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Indique un código de Usuario y Correo Electrónico registrado en el sistema"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Usuario 
         Height          =   312
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   4452
         _Version        =   1245185
         _ExtentX        =   7853
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Email 
         Height          =   312
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   4452
         _Version        =   1245185
         _ExtentX        =   7853
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnToken_Envia 
         Height          =   492
         Left            =   5160
         TabIndex        =   6
         Top             =   1320
         Width           =   1812
         _Version        =   1245185
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Enviar TOKEN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   2
         Picture         =   "frmLogon_Renueva.frx":096B
         ImageAlignment  =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   2052
      End
   End
   Begin XtremeSuiteControls.RadioButton RbtToken 
      Height          =   852
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2052
      _Version        =   1245185
      _ExtentX        =   3619
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Ya Cuento con el TOKEN y quiero renovar mi contraseña"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      Appearance      =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TOKEN digital para Renovar su Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8292
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmLogon_Renueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String

Private Sub btnToken_Activa_Click()
Call sbValida_Token

If Len(vMensaje) = 0 Then
    glogon.Usuario = FlatEdit_Token_Usuario.Text
    frmLogon_Renueva_Clave.Show vbModal
    Unload Me
Else
    MsgBox vMensaje, vbExclamation
End If
End Sub

Private Sub btnToken_Envia_Click()

Call sbValida_Datos

If Len(vMensaje) = 0 Then
    Call sbToken_Envia
Else
    MsgBox vMensaje, vbExclamation
End If
    
End Sub

Private Sub Form_Load()
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 Call PGX_Portal_Access
End Sub

Private Sub RbtToken_Click(Index As Integer)

 If Index = 0 Then
    GroupBox_Solicita.Visible = True
    GroupBox_Activa.Visible = False
 Else
    GroupBox_Solicita.Visible = False
    GroupBox_Activa.Visible = True
 End If
 
End Sub

Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function


Private Function fxToken_Genera() As String
Dim i As Integer, Largo As Integer, Cadena As String, x As Integer

 Cadena = ""
 Largo = Aleatorio(8, 16)
 For x = 1 To Largo
        i = Aleatorio(40, 90)
        Cadena = Cadena & Chr(i)
 Next x
 
 fxToken_Genera = Cadena
End Function


Private Sub sbValida_Datos()
Dim strSQL As String, rs As New ADODB.Recordset

vMensaje = ""

strSQL = "select dbo.fxSEG_Token_Valida_Usuario('" & FlatEdit_Usuario.Text & "','" & FlatEdit_Email.Text & "','') as 'Resultado'"
Call OpenRecordSet(rs, strSQL, 1)
If rs!Resultado = 0 Then
    vMensaje = "Datos del Usuario y Correo no Localizados!"
End If

End Sub

Private Sub sbValida_Token()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTokenKey As String

vMensaje = ""
vTokenKey = SIFGlobal.fxStringCifrado(FlatEdit_Token.Text)

strSQL = "select dbo.fxSEG_Token_Valida('" & FlatEdit_Token_Usuario.Text & "','" & vTokenKey & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL, 1)
Select Case rs!Resultado
    Case 1 'Ok
    Case 0 'No Existe
        vMensaje = "Token no es válido!"
    Case -1 'Está vencido
        vMensaje = "El Token suministrado se encuentra vencido!"
End Select

rs.Close

End Sub


Private Sub sbToken_Envia()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vToken As String, vTokenKey As String

On Error GoTo vError

vToken = fxToken_Genera
vTokenKey = SIFGlobal.fxStringCifrado(vToken)

strSQL = "exec spSEG_Token_Registra '" & FlatEdit_Usuario.Text & "','" & vToken & "','" & vTokenKey & "'"
Call ConectionExecute(strSQL, 1)

FlatEdit_Usuario.Text = ""
FlatEdit_Email.Text = ""

MsgBox "Se ha enviado un TOKEN para renovación de su contraseña a su correo electrónico. Favor revisar su correo e ingresa a la opción de Ingreso del Token", vbInformation

Call RbtToken_Click(1)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

